# -*- coding: utf-8 -*-
"""
使用方式:
  from precision_engine_v5 import (
      boost_process_priority,
      boost_thread_priority,
      set_thread_affinity,
      precise_sleep_v5,
      auto_warmup,
  )
"""

import ctypes
from ctypes import wintypes
import time
import threading
import struct
import atexit

# ==============================================================================
# 0. WinAPI
# ==============================================================================
kernel32 = ctypes.windll.kernel32
winmm = ctypes.windll.winmm

winmm.timeBeginPeriod(1)

def _cleanup():
    try:
        winmm.timeEndPeriod(1)
    except Exception:
        pass
atexit.register(_cleanup)

_is_64bit = struct.calcsize("P") == 8

# ==============================================================================
# 1. 高分辨率定时器 (windll — 释放GIL)
# ==============================================================================
CREATE_WAITABLE_TIMER_HIGH_RESOLUTION = 0x00000002
TIMER_ALL_ACCESS = 0x1F0003


class HighResTimer:
    __slots__ = ('handle', '_is_high_res')

    def __init__(self):
        self._is_high_res = False
        try:
            kernel32.CreateWaitableTimerExW.restype = wintypes.HANDLE
            kernel32.CreateWaitableTimerExW.argtypes = [
                wintypes.LPVOID, wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
            ]
            self.handle = kernel32.CreateWaitableTimerExW(
                None, None, CREATE_WAITABLE_TIMER_HIGH_RESOLUTION, TIMER_ALL_ACCESS
            )
            if self.handle and self.handle != -1:
                self._is_high_res = True
            else:
                self.handle = kernel32.CreateWaitableTimerExW(None, None, 0, TIMER_ALL_ACCESS)
        except Exception:
            self.handle = kernel32.CreateWaitableTimerW(None, True, None)

    @property
    def is_high_resolution(self):
        return self._is_high_res

    def wait(self, seconds):
        if seconds <= 0:
            return 0.0
        due_time_100ns = ctypes.c_longlong(int(-seconds * 10_000_000))
        kernel32.SetWaitableTimer(self.handle, ctypes.byref(due_time_100ns), 0, None, None, False)
        kernel32.WaitForSingleObject(self.handle, 0xFFFFFFFF)
        return 0.0

    def close(self):
        if self.handle:
            kernel32.CloseHandle(self.handle)
            self.handle = None


_timer_local = threading.local()


def _get_timer():
    timer = getattr(_timer_local, 'timer', None)
    if timer is None:
        timer = HighResTimer()
        _timer_local.timer = timer
    return timer


# ==============================================================================
# 2. QPC 基础设施
# ==============================================================================
class LARGE_INTEGER(ctypes.Structure):
    _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]


kernel32.QueryPerformanceCounter.argtypes = [ctypes.POINTER(LARGE_INTEGER)]
kernel32.QueryPerformanceCounter.restype = wintypes.BOOL
kernel32.QueryPerformanceFrequency.argtypes = [ctypes.POINTER(LARGE_INTEGER)]
kernel32.QueryPerformanceFrequency.restype = wintypes.BOOL

_freq = LARGE_INTEGER()
kernel32.QueryPerformanceFrequency(ctypes.byref(_freq))
QPC_FREQ = _freq.LowPart | (_freq.HighPart << 32)

# ==============================================================================
# 3. 原生 x64 机器码忙等
# ==============================================================================
_native_spin_func = None
if _is_64bit:
    _qpc_func_ptr = ctypes.cast(kernel32.QueryPerformanceCounter, ctypes.c_void_p).value

    # 机器码忙等函数
    _native_mc = bytes([
        0x53, 0x41, 0x54, 0x48, 0x83, 0xEC, 0x28,
        0x48, 0x89, 0xCB, 0x49, 0x89, 0xD4,
        0x48, 0x8D, 0x4C, 0x24, 0x20, 0x41, 0xFF, 0xD4,
        0x48, 0x8B, 0x44, 0x24, 0x20, 0x48, 0x39, 0xD8,
        0x72, 0xEE,
        0x48, 0x83, 0xC4, 0x28, 0x41, 0x5C, 0x5B, 0xC3,
    ])

    def _create_native_spin_func():
        kernel32.VirtualAlloc.restype = ctypes.c_void_p
        kernel32.VirtualAlloc.argtypes = [
            ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD
        ]
        buf = kernel32.VirtualAlloc(None, 4096, 0x1000, 0x40)
        if not buf:
            return None
        ctypes.memmove(buf, _native_mc, len(_native_mc))
        return ctypes.CFUNCTYPE(None, ctypes.c_uint64, ctypes.c_void_p)(buf)

    _native_spin_func = _create_native_spin_func()

# ==============================================================================
# 4. 混合忙等 (100us 脉冲 + Python yield)
# ==============================================================================
_PULSE_DURATION = 0.0001  # 100us
_native_spin_func_ref = _native_spin_func


def _hybrid_spin_wait(end_time_perf, stop_event, perf_counter):
    spin_func = _native_spin_func_ref
    if spin_func is not None:
        freq = QPC_FREQ
        while perf_counter() < end_time_perf:
            if stop_event.is_set():
                return
            now = perf_counter()
            target_qpc = int(min(now + _PULSE_DURATION, end_time_perf) * freq)
            spin_func(target_qpc, _qpc_func_ptr)
    else:
        while perf_counter() < end_time_perf:
            if stop_event.is_set():
                return
            time.sleep(0)


# ==============================================================================
# 5. 自适应偏差补偿器 (v5.1 修复: 作用于最终目标)
# ==============================================================================
class AdaptiveCompensator:
    """
    v5.1: bias -> 直接调整 最终目标 (总延迟居中到 target)
    """
    __slots__ = ('bias', 'alpha', 'warmup', 'count', '_errors', '_window', '_mean', '_var_sum')

    def __init__(self, alpha=0.04, warmup=30):
        self.bias = 0.0
        self.alpha = alpha
        self.warmup = warmup
        self.count = 0
        self._errors = []
        self._window = 60
        self._mean = 0.0
        self._var_sum = 0.0

    @property
    def is_converged(self):
        return self.count >= self.warmup

    def _is_outlier(self, value):
        n = len(self._errors)
        if n < 20:
            return False
        mean = self._mean
        std = (self._var_sum / n) ** 0.5
        return std > 0 and abs(value - mean) > 3 * std

    def get_adjusted_target(self, original_target, min_total=0.0002):
        if self.count < self.warmup:
            return original_target
        adjusted = max(original_target - self.bias, min_total)
        return max(adjusted, original_target * 0.1)

    def update(self, actual_delay, target):
        error = actual_delay - target
        self.count += 1
        self._errors.append(error)

        n = len(self._errors)
        if n > self._window:
            old_err = self._errors.pop(0)
            # 增量更新均值和方差和
            old_mean = self._mean
            self._mean += (error - old_err - self._mean) / n
            self._var_sum += (error - old_mean) * (error - old_mean) - \
                             (old_err - old_mean) * (old_err - old_mean)
        else:
            old_mean = self._mean
            self._mean += (error - old_mean) / n
            if n > 1:
                self._var_sum += (error - old_mean) * (error - self._mean)

        if self._is_outlier(error):
            return
        self.bias += self.alpha * error
        self.bias = max(-0.0002, min(0.002, self.bias))

    def reset(self):
        self.bias = 0.0
        self.count = 0
        self._errors.clear()
        self._mean = 0.0
        self._var_sum = 0.0


# ==============================================================================
# 6. 一次标定: 混合忙等固有开销
# ==============================================================================
_HYBRID_SPIN_OVERHEAD = None


def _calibrate_hybrid_spin_overhead():
    global _HYBRID_SPIN_OVERHEAD
    _get_timer().wait(0.001)
    samples = []
    stop = threading.Event()
    pc = time.perf_counter
    for _ in range(300):
        start = pc()
        _hybrid_spin_wait(start + 0.000001, stop, pc)
        samples.append((pc() - start) * 1_000_000)
    samples.sort()
    _HYBRID_SPIN_OVERHEAD = samples[len(samples) // 2] / 1_000_000
    return _HYBRID_SPIN_OVERHEAD


# ==============================================================================
# 7. 线程本地状态 + 优先级管理
# ==============================================================================
_comp_local = threading.local()


def _get_compensator():
    comp = getattr(_comp_local, 'comp', None)
    if comp is None:
        comp = AdaptiveCompensator(alpha=0.04, warmup=30)
        _comp_local.comp = comp
    return comp


def boost_process_priority():
    kernel32.SetPriorityClass(kernel32.GetCurrentProcess(), 0x00000080)


def boost_thread_priority():
    kernel32.SetThreadPriority(kernel32.GetCurrentThread(), 15)


def set_thread_affinity(core_index=3):
    """默认绑定到第4个核心，避免与系统线程竞争"""
    kernel32.SetThreadAffinityMask(kernel32.GetCurrentThread(), 1 << core_index)


def auto_warmup(target_s=0.100, iterations=30):
    stop = threading.Event()
    comp = _get_compensator()
    for _ in range(iterations):
        precise_sleep_v5(target_s, stop, comp)


# ==============================================================================
# 8. 核心: precise_sleep_v5
# ==============================================================================
_TIMER_LATE_THRESHOLD = 0.0008  # 800us
_TIMER = _get_timer


def precise_sleep_v5(seconds, stop_event, compensator=None):
    if seconds <= 0 or stop_event.is_set():
        return

    if compensator is None:
        compensator = _get_compensator()

    global _HYBRID_SPIN_OVERHEAD
    if _HYBRID_SPIN_OVERHEAD is None:
        _calibrate_hybrid_spin_overhead()

    timer = _get_timer()
    perf_counter = time.perf_counter
    start_perf = perf_counter()

    # ★ 关键修复: bias 作用于最终目标
    adjusted_target = compensator.get_adjusted_target(seconds)
    actual_end_perf = start_perf + adjusted_target

    # 自适应忙等裕度（简化为两档）
    spin_margin = 0.00015 if adjusted_target < 0.005 else 0.0002

    # 极短延迟: 直接忙等
    if adjusted_target <= spin_margin * 2:
        _hybrid_spin_wait(actual_end_perf, stop_event, perf_counter)
        compensator.update(perf_counter() - start_perf, adjusted_target)
        return

    intermediate_end_perf = actual_end_perf - spin_margin

    # ---- Phase 1: 粗睡眠 (windll, 释放GIL) ----
    coarse_threshold, max_chunk = (0.005, 0.005) if seconds > 0.05 else \
                                  (0.003, 0.003) if seconds > 0.01 else (0.001, 0.002)

    while perf_counter() < intermediate_end_perf - coarse_threshold:
        if stop_event.is_set():
            compensator.update(perf_counter() - start_perf, seconds)
            return
        remaining = intermediate_end_perf - perf_counter()
        chunk = min(max_chunk, remaining - coarse_threshold * 0.5)
        if chunk > 0.0001:
            before_wait = perf_counter()
            timer.wait(chunk)
            after_wait = perf_counter()
            # ★ 迟到检测
            if (after_wait - before_wait - chunk) > _TIMER_LATE_THRESHOLD:
                break

    # ---- Phase 2: 中精度 (windll, 释放GIL) ----
    fine_threshold = 0.0002
    while perf_counter() < intermediate_end_perf - fine_threshold:
        if stop_event.is_set():
            compensator.update(perf_counter() - start_perf, seconds)
            return
        remaining = intermediate_end_perf - perf_counter()
        chunk = min(0.0001, remaining - fine_threshold * 0.5)
        if chunk > 0.00001:
            timer.wait(chunk)

    # ---- Phase 3: 混合忙等 ----
    if not stop_event.is_set():
        _hybrid_spin_wait(actual_end_perf, stop_event, perf_counter)

    compensator.update(perf_counter() - start_perf, seconds)


# ==============================================================================
# 9. 基准测试
# ==============================================================================
def benchmark_precision(iterations=500, warmup=50):
    import statistics

    boost_process_priority()
    boost_thread_priority()
    set_thread_affinity(core_index=1)

    timer = _get_timer()
    comp = _get_compensator()

    print("=" * 70)
    print(" Tier 0 精度引擎 v5.1 基准测试")
    print(f" 高分辨率定时器: {'TSC' if timer.is_high_resolution else '回退'}")
    spin_type = "x64 混合(100us)" if _native_spin_func else "Python回退"
    print(f" 忙等模式: {spin_type}")
    print(f" GIL策略: windll释放 + 混合yield(每100us)")
    print(f" 全局时钟: 1ms (timeBeginPeriod)")
    print(f" 补偿策略: 最终目标补偿 (居中)")
    if _HYBRID_SPIN_OVERHEAD is not None:
        print(f" 混合忙等开销标定: {_HYBRID_SPIN_OVERHEAD * 1e6:.1f}us")
    print(f" 测试: 预热{warmup}次 + 测量{iterations}次")
    print("=" * 70)

    test_cases = [("1ms", 0.001), ("5ms", 0.005), ("10ms", 0.010),
                  ("50ms", 0.050), ("100ms", 0.100)]

    for name, target_s in test_cases:
        stop = threading.Event()
        comp.reset()

        for _ in range(warmup):
            precise_sleep_v5(target_s, stop, comp)

        target_ms = target_s * 1000
        samples = [(time.perf_counter() - time.perf_counter()) * 0]  # placeholder for type
        s = time.perf_counter
        samples = [(s() - s()) * 0]  # reset properly
        samples = []
        for _ in range(iterations):
            st = time.perf_counter()
            precise_sleep_v5(target_s, stop, comp)
            samples.append((time.perf_counter() - st) * 1000)

        mean = statistics.mean(samples)
        median = statistics.median(samples)
        stdev = statistics.stdev(samples)
        max_err = max(abs(v - target_ms) for v in samples)
        outliers = sum(1 for v in samples if abs(v - target_ms) > 0.1)

        non_outliers = [v for v in samples if abs(v - target_ms) <= 0.1]
        if len(non_outliers) > 1:
            nm = statistics.mean(non_outliers)
            nmed = statistics.median(non_outliers)
            nsd = statistics.stdev(non_outliers)
        else:
            nm = nmed = nsd = float("nan")

        below = sum(1 for v in samples if v < target_ms)
        above = len(samples) - below

        print(f"\n  {name:>5s} 目标: {target_ms:>7.2f}ms")
        print(f"  范围: [{min(samples):.4f}, {max(samples):.4f}]")
        print(f"  均值: {mean:.4f}ms (偏移{(mean - target_ms) * 1000:+.1f}us)")
        print(f"  中位数: {median:.4f}ms (偏移{(median - target_ms) * 1000:+.1f}us)")
        print(f"  最大误差: {max_err:.4f}ms  异常值(>100us): {outliers}/{iterations}")
        print(f"  下方/上方: {below}/{above} ({below / len(samples) * 100:.0f}%/{above / len(samples) * 100:.0f}%)")
        print(f"  收敛偏差: {comp.bias * 1e6:.1f}us")
        print(f"  非异常值: 均值={nm:.4f}ms 中位数={nmed:.4f}ms sigma={nsd:.4f}ms")

    print("\n" + "=" * 70)


if __name__ == "__main__":
    benchmark_precision()
