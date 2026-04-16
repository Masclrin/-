# -*- coding: utf-8 -*-
"""
框架精度对比测试脚本
===================
用法:
  python macro_precision_benchmark.py

功能:
  1. 自动生成多组标准测试宏（不同延迟梯度）
  2. 可切换 PI参数（1.9 / 1.91 参数）及其它优化参数配置进行对比
  3. 统计: 偏移、抖动、累计偏差、尾部延迟、CPU开销
  4. 输出表格化对比报告 + CSV 导出
  5. 可选: 模拟 CPU 负载下测试鲁棒性
"""

import os
import sys
import json
import time
import gc
import threading
import ctypes
import statistics
import csv
from datetime import datetime

# ── 导入框架依赖 ──
# 注意: 请确保 precision_engine_v5.py 在同目录或 Python 路径中
from precision_engine_v5 import (
    boost_process_priority,
    boost_thread_priority,
    set_thread_affinity,
    precise_sleep_v5,
    auto_warmup,
)

# ==============================================================================
# 测试配置
# ==============================================================================

# 测试延迟梯度 (毫秒)
TEST_DELAY_MS = [5, 50]

# 每组测试重复次数
TEST_REPEATS = 5

# 每次执行中的步数 (模拟短宏 vs 长宏)
TEST_STEPS_PER_RUN = [5, 40]

# 是否输出每步明细 (用于调试，关闭可加速)
VERBOSE_STEP_LOG = False

# 是否模拟 CPU 负载 (后台占一个核心)
SIMULATE_CPU_LOAD = False
CPU_LOAD_THREADS = 2

# 结果输出目录
OUTPUT_DIR = "benchmark_results"

# CSV 导出
EXPORT_CSV = True

# 输入类型维度: 键盘与鼠标分别测试
TEST_ACTION_TYPES = ["keyboard", "mouse"]

# 动作发包后端: interception(真实驱动发包) / simulated(旧模拟模式)
ACTION_BACKEND = "interception"

# Interception 基准发包参数
# 键盘使用较少业务副作用的扫描码进行 down+up 配对发送
TEST_KEY_SCAN_CODE = 0x7F
# 鼠标默认右键，降低误触业务交互风险；如需测左键改为 True
TEST_MOUSE_LEFT = False
INTERCEPTION_WARMUP_SENDS = 30

# ── 需要测试的配置方案 ──
# 每个 dict 代表一套 PI 控制器 + 框架参数
CONFIG_PRESETS1 = {
    "v1.9_保守": {
        "lag_enabled": True,
        "lag_error_trigger_ms": 1.0,
        "lag_error_target_ms": 0.4,
        "lag_kp": 0.35,
        "lag_ki": 0.01,
        "lag_integral_max": 2.0,
        "lag_integral_decay": 0.998,
        "lag_max_step_comp_pct": 0.01,
        "lag_min_step_delay_ms": 0.3,
        "min_sleep_ms": 1,         # max(0.001, sleep_time)
        "use_pro_audio": False,
        "use_realtime": False,
        "use_nt_timer_05ms": False,
    },
    "v1.91_激进": {
        "lag_enabled": True,
        "lag_error_trigger_ms": 0.0,
        "lag_error_target_ms": 0.0,
        "lag_kp": 0.8,
        "lag_ki": 0.05,
        "lag_integral_max": 5.0,
        "lag_integral_decay": 0.998,
        "lag_max_step_comp_pct": 0.15,
        "lag_min_step_delay_ms": 0.0,
        "min_sleep_ms": 0.0,          # 允许 0ms
        "use_pro_audio": True,
        "use_realtime": True,
        "use_nt_timer_05ms": True,
    },
    "保守参数_无补偿": {
        "lag_enabled": False,
        "min_sleep_ms": 1.0,
        "use_pro_audio": False,
        "use_realtime": False,
        "use_nt_timer_05ms": False,
    },
    "仅最低睡眠_无补偿": {
        "lag_enabled": False,
        "min_sleep_ms": 0.0,
        "use_pro_audio": False,
        "use_realtime": False,
        "use_nt_timer_05ms": False,
    },
    "仅调整至实时_无补偿": {
        "lag_enabled": False,
        "min_sleep_ms": 1.0,
        "use_pro_audio": False,
        "use_realtime": True,
        "use_nt_timer_05ms": False,
    },
    "仅NTTimer0.5ms_无补偿": {
        "lag_enabled": False,
        "min_sleep_ms": 1.0,
        "use_pro_audio": False,
        "use_realtime": False,
        "use_nt_timer_05ms": True,
    },
    "NTTimer0.5ms_+实时": {
        "lag_enabled": False,
        "min_sleep_ms": 1.0,
        "use_pro_audio": False,
        "use_realtime": True,
        "use_nt_timer_05ms": True,
    },
    "仅ProAudio_无补偿": {
        "lag_enabled": False,
        "min_sleep_ms": 1.0,
        "use_pro_audio": True,
        "use_realtime": False,
        "use_nt_timer_05ms": False,
    },
    "全开，不PI": {
        "lag_enabled": False,
        "min_sleep_ms": 0.0,
        "use_pro_audio": True,
        "use_realtime": True,
        "use_nt_timer_05ms": True,
    },
    "仅PI": {
        "lag_enabled": True,
        "lag_error_trigger_ms": 0.0,
        "lag_error_target_ms": 0.0,
        "lag_kp": 0.8,
        "lag_ki": 0.05,
        "lag_integral_max": 5.0,
        "lag_integral_decay": 0.998,
        "lag_max_step_comp_pct": 0.15,
        "lag_min_step_delay_ms": 0.0,
        "min_sleep_ms": 1.0,
        "use_pro_audio": False,
        "use_realtime": False,
        "use_nt_timer_05ms": False,
    },
}

CONFIG_PRESETS2 = {    
    "NTTimer0.5ms_+实时": {
        "lag_enabled": False,
        "min_sleep_ms": 1.0,
        "use_pro_audio": False,
        "use_realtime": True,
        "use_nt_timer_05ms": True,
    },
    }  # 可根据需要选择不同的预设组合进行测试

# 当前启用的配置集合: 按需在 CONFIG_PRESETS1 / CONFIG_PRESETS2 间切换
CONFIG_PRESETS = CONFIG_PRESETS1


# ==============================================================================
# PI 累计误差控制器 (从框架提取，独立可控)
# ==============================================================================

class CumulativeLagController:
    """PI 累计误差控制器 — 参数可外部注入"""

    def __init__(self, cfg):
        self.kp = cfg["lag_kp"]
        self.ki = cfg["lag_ki"]
        self.error_trigger = cfg["lag_error_trigger_ms"] / 1000.0
        self.error_target = cfg["lag_error_target_ms"] / 1000.0
        self.integral_max = cfg["lag_integral_max"]
        self.integral_decay = cfg["lag_integral_decay"]
        self.max_step_comp_pct = cfg["lag_max_step_comp_pct"]
        self.min_step_delay = cfg["lag_min_step_delay_ms"] / 1000.0
        # Kalman
        self._est = 0.0
        self._cov = 1.0
        self._q = 1e-6
        self._r = 2e-5
        # PI integral
        self._integral = 0.0
        # Stats
        self.total_compensated = 0.0
        self.total_recovered = 0.0
        self.compensation_count = 0

    def _kalman_update(self, measurement):
        self._cov += self._q
        gain = self._cov / (self._cov + self._r)
        self._est += gain * (measurement - self._est)
        self._cov *= (1.0 - gain)
        return self._est

    def compute_compensation(self, current_time, ideal_time, next_delay, action_type=""):
        raw_lag = current_time - ideal_time
        smoothed_lag = self._kalman_update(raw_lag)
        if action_type == "wait":
            return 0.0
        if abs(smoothed_lag) <= self.error_trigger:
            self._integral *= self.integral_decay
            return 0.0
        error_excess = smoothed_lag - self.error_target
        p_term = self.kp * error_excess
        self._integral += error_excess
        self._integral = max(-self.integral_max, min(self._integral, self.integral_max))
        i_term = self.ki * self._integral
        total_comp = p_term + i_term
        max_by_pct = next_delay * self.max_step_comp_pct
        compensation = max(-max_by_pct, min(total_comp, max_by_pct))
        if compensation > 0:
            self.total_compensated += compensation
        elif compensation < 0:
            self.total_recovered += abs(compensation)
        self.compensation_count += 1
        return compensation


# ==============================================================================
# 进程/线程提权辅助
# ==============================================================================

kernel32 = ctypes.windll.kernel32

# Interception 运行时状态
_interception_ctx = None
_interception_key_down = None
_interception_key_up = None

def apply_realtime_priority(enable_realtime=True):
    """按开关设置进程优先级: True=REALTIME, False=NORMAL。"""
    try:
        handle = kernel32.OpenProcess(0x1F0FFF, False, kernel32.GetCurrentProcessId())
        target_class = 0x00000100 if enable_realtime else 0x00000020  # REALTIME / NORMAL
        ok = bool(kernel32.SetPriorityClass(handle, target_class))
        kernel32.CloseHandle(handle)
        return ok
    except Exception:
        return False


def apply_nt_timer_05ms(enable_05ms=True):
    """按开关请求/释放 0.5ms 系统计时器分辨率。"""
    try:
        current = ctypes.c_ulong(0)
        status = ctypes.windll.ntdll.NtSetTimerResolution(5000, bool(enable_05ms), ctypes.byref(current))
        return status == 0
    except Exception:
        return False


def apply_runtime_tuning(cfg):
    """根据配置应用运行时系统参数。"""
    use_realtime = bool(cfg.get("use_realtime", False))
    use_nt_timer = bool(cfg.get("use_nt_timer_05ms", False))

    apply_realtime_priority(use_realtime)
    apply_nt_timer_05ms(use_nt_timer)

def apply_pro_audio():
    """MMCSS Pro Audio 注入"""
    try:
        avrt = ctypes.windll.avrt
        task_index = ctypes.c_ulong(0)
        h_task = avrt.AvSetMmThreadCharacteristicsW("Pro Audio", ctypes.byref(task_index))
        if h_task:
            avrt.AvSetMmThreadPriority(h_task, 2)
            return True
        return False
    except Exception:
        return False


def init_action_backend():
    """初始化动作发包后端。默认使用 Interception 真实发包。"""
    global _interception_ctx, _interception_key_down, _interception_key_up

    if ACTION_BACKEND != "interception":
        print("[输入] 动作发包后端: simulated (仅模拟调用开销)")
        return True

    try:
        from interception_input import (
            get_interception_context,
            INTERCEPTION_KEY_DOWN,
            INTERCEPTION_KEY_UP,
        )

        _interception_ctx = get_interception_context()
        _interception_key_down = INTERCEPTION_KEY_DOWN
        _interception_key_up = INTERCEPTION_KEY_UP

        # 预热驱动路径，避免首次发包初始化放大统计波动
        for _ in range(INTERCEPTION_WARMUP_SENDS):
            _interception_ctx.send_key(TEST_KEY_SCAN_CODE, _interception_key_up, 0)
            _interception_ctx.send_mouse_click(left=TEST_MOUSE_LEFT, up=True)

        print("[输入] 动作发包后端: interception (真实驱动发包)")
        return True
    except Exception as e:
        print(f"[输入] Interception 初始化失败: {e}")
        return False


def dispatch_action(action_type, step_index):
    """执行一次动作发包。默认发送 down+up 配对，计入真实驱动调用开销。"""
    if ACTION_BACKEND == "interception":
        if _interception_ctx is None:
            raise RuntimeError("Interception 未初始化")

        if action_type == "mouse":
            _interception_ctx.send_mouse_click(left=TEST_MOUSE_LEFT, up=False)
            _interception_ctx.send_mouse_click(left=TEST_MOUSE_LEFT, up=True)
        else:
            _interception_ctx.send_key(TEST_KEY_SCAN_CODE, _interception_key_down, 0)
            _interception_ctx.send_key(TEST_KEY_SCAN_CODE, _interception_key_up, 0)
        return

    # 回退: 旧模拟路径
    if action_type == "mouse":
        _dummy = ctypes.c_long(step_index)
        _dummy.value ^= 0x55
    else:
        _dummy = ctypes.c_int(step_index)
        _dummy.value ^= 0xAA


# ==============================================================================
# 核心测试循环 (模拟宏执行)
# ==============================================================================

def run_single_test(delay_ms, steps, cfg, stop_event, action_type="keyboard"):
    """
    模拟执行一个 steps 步、每步 delay_ms 的宏。
    返回详细的每步统计数据。
    """
    delay_s = delay_ms / 1000.0
    min_sleep_s = cfg.get("min_sleep_ms", 1.0) / 1000.0

    # 初始化控制器
    lag_ctrl = CumulativeLagController(cfg) if cfg.get("lag_enabled") else None

    step_records = []
    start_time = time.perf_counter()
    ideal_time = start_time

    gc_was_enabled = gc.isenabled()
    if gc_was_enabled:
        gc.disable()

    try:
        for i in range(steps):
            if stop_event.is_set():
                break

            # 动作执行: 默认走 Interception 真实驱动发包路径
            action_start = time.perf_counter()
            dispatch_action(action_type, i)
            action_end = time.perf_counter()
            action_cost = action_end - action_start

            current_time = action_end

            # PI 补偿
            if lag_ctrl is not None:
                compensation = lag_ctrl.compute_compensation(
                    current_time, ideal_time, delay_s, "action"
                )
            else:
                compensation = 0.0

            # 推进理想时间线
            ideal_time += delay_s

            # 实际睡眠
            sleep_time = delay_s - compensation
            sleep_time = max(min_sleep_s, sleep_time) if min_sleep_s > 0 else max(0.0, sleep_time)

            sleep_start = time.perf_counter()
            if sleep_time > 0:
                precise_sleep_v5(sleep_time, stop_event)
            actual_wait = time.perf_counter() - sleep_start

            # 统计
            step_total = action_cost + actual_wait
            step_error = step_total - delay_s
            cum_bias = (time.perf_counter() - ideal_time) * 1000.0  # ms

            step_records.append({
                "step": i + 1,
                "action_type": action_type,
                "action_cost_ms": action_cost * 1000.0,
                "sleep_target_ms": (delay_s - compensation) * 1000.0,
                "actual_wait_ms": actual_wait * 1000.0,
                "step_error_ms": step_error * 1000.0,
                "cum_bias_ms": cum_bias,
                "compensation_ms": compensation * 1000.0,
            })
    finally:
        if gc_was_enabled:
            gc.enable()

    total_elapsed = time.perf_counter() - start_time
    expected_total = delay_s * steps

    return {
        "action_type": action_type,
        "steps": step_records,
        "total_elapsed_s": total_elapsed,
        "expected_total_s": expected_total,
        "total_error_ms": (total_elapsed - expected_total) * 1000.0,
        "lag_stats": {
            "compensated_ms": lag_ctrl.total_compensated * 1000.0 if lag_ctrl else 0,
            "recovered_ms": lag_ctrl.total_recovered * 1000.0 if lag_ctrl else 0,
            "comp_count": lag_ctrl.compensation_count if lag_ctrl else 0,
        } if lag_ctrl else None,
    }


def compute_stats(step_records):
    """从每步记录计算统计指标"""
    if not step_records:
        return {}

    errors = [r["step_error_ms"] for r in step_records]
    waits = [r["actual_wait_ms"] for r in step_records]
    costs = [r["action_cost_ms"] for r in step_records]
    biases = [r["cum_bias_ms"] for r in step_records]

    return {
        "mean_error_ms": statistics.mean(errors),
        "median_error_ms": statistics.median(errors),
        "std_error_ms": statistics.stdev(errors) if len(errors) > 1 else 0,
        "max_error_ms": max(errors),
        "min_error_ms": min(errors),
        "p99_error_ms": sorted(errors)[int(len(errors) * 0.99)] if len(errors) > 10 else max(errors),
        "p95_error_ms": sorted(errors)[int(len(errors) * 0.95)] if len(errors) > 10 else max(errors),
        "mean_wait_ms": statistics.mean(waits),
        "mean_action_cost_ms": statistics.mean(costs),
        "final_cum_bias_ms": biases[-1],
        "max_cum_bias_ms": max(abs(b) for b in biases),
        "abs_mean_bias_ms": statistics.mean(abs(b) for b in biases),
    }


# ==============================================================================
# CPU 负载模拟器
# ==============================================================================

_cpu_load_stop = threading.Event()

def _cpu_load_worker():
    """占满一个核心的忙等循环"""
    while not _cpu_load_stop.is_set():
        pass

def start_cpu_load(threads=2):
    _cpu_load_stop.clear()
    workers = []
    for _ in range(threads):
        t = threading.Thread(target=_cpu_load_worker, daemon=True)
        t.start()
        workers.append(t)
    return workers

def stop_cpu_load():
    _cpu_load_stop.set()


# ==============================================================================
# 主测试流程
# ==============================================================================

def run_benchmark():
    """执行完整基准测试"""

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    auto_warmup()
    print("=" * 80)
    print("  宏执行框架 精度对比基准测试")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    # 系统信息
    print(f"\n[系统] Python {sys.version}")
    print(f"[系统] CPU 核心数: {os.cpu_count()}")
    print(f"[系统] 测试延迟梯度: {TEST_DELAY_MS} ms")
    print(f"[系统] 每组重复: {TEST_REPEATS} 次")
    print(f"[系统] 步数梯度: {TEST_STEPS_PER_RUN}")
    print(f"[系统] 输入类型: {TEST_ACTION_TYPES}")
    print(f"[系统] 动作发包后端: {ACTION_BACKEND}")
    print(f"[系统] CPU 负载模拟: {'开启 ({}线程)'.format(CPU_LOAD_THREADS) if SIMULATE_CPU_LOAD else '关闭'}")
    print(f"[系统] 配置方案: {list(CONFIG_PRESETS.keys())}")

    # 全局提权
    boost_process_priority()

    # 初始化动作发包后端
    if not init_action_backend():
        raise RuntimeError("动作发包后端初始化失败。请确认 Interception 驱动可用，或将 ACTION_BACKEND 改为 simulated")

    # CPU 负载
    if SIMULATE_CPU_LOAD:
        print("\n[负载] 启动 CPU 负载模拟...")
        load_workers = start_cpu_load(CPU_LOAD_THREADS)
        time.sleep(1)

    # 存储所有结果
    all_results = {}  # (preset_name, action_type, delay_ms, steps) -> [run_results...]

    total_tests = len(CONFIG_PRESETS) * len(TEST_ACTION_TYPES) * len(TEST_DELAY_MS) * len(TEST_STEPS_PER_RUN) * TEST_REPEATS
    completed = 0

    try:
        for preset_name, cfg in CONFIG_PRESETS.items():
            print(f"\n{'─' * 70}")
            print(f"  ▶ 配置方案: {preset_name}")
            if cfg.get("lag_enabled"):
                print(f"    PI: Kp={cfg.get('lag_kp', 'N/A')} Ki={cfg.get('lag_ki', 'N/A')} "
                      f"trigger={cfg.get('lag_error_trigger_ms', 'N/A')}ms "
                      f"target={cfg.get('lag_error_target_ms', 'N/A')}ms")
            else:
                print("    PI: 关闭")
            max_comp_pct = cfg.get("lag_max_step_comp_pct")
            max_comp_text = f"{max_comp_pct * 100:.1f}%" if isinstance(max_comp_pct, (int, float)) else "N/A"
            print(f"    补偿上限: {max_comp_text} "
                f"min_sleep: {cfg.get('min_sleep_ms', 'N/A')}ms")
            print(f"    ProAudio: {cfg.get('use_pro_audio')} "
                  f"Realtime: {cfg.get('use_realtime')} "
                  f"0.5ms时钟: {cfg.get('use_nt_timer_05ms')}")
            print(f"{'─' * 70}")

            for action_type in TEST_ACTION_TYPES:
                print(f"    输入类型: {action_type}")
                for delay_ms in TEST_DELAY_MS:
                    for steps in TEST_STEPS_PER_RUN:
                        run_stats_list = []

                        for run_idx in range(TEST_REPEATS):
                            completed += 1
                            pct = completed / total_tests * 100
                            print(f"\r  [{pct:5.1f}%] {preset_name} | {action_type} | "
                                  f"delay={delay_ms:>3d}ms | steps={steps:>3d} | "
                                  f"run={run_idx + 1}/{TEST_REPEATS}", end="", flush=True)

                            stop_event = threading.Event()

                            # 线程级设置 (每次运行新线程以干净状态)
                            result_holder = [None]

                            def _worker():
                                boost_thread_priority()
                                total_cores = os.cpu_count() or 4
                                set_thread_affinity(total_cores - 2)
                                apply_runtime_tuning(cfg)
                                if cfg.get("use_pro_audio"):
                                    apply_pro_audio()
                                result_holder[0] = run_single_test(delay_ms, steps, cfg, stop_event, action_type)

                            t = threading.Thread(target=_worker)
                            t.start()
                            t.join(timeout=60)

                            if result_holder[0] is None:
                                print(" [超时!]")
                                continue

                            result = result_holder[0]
                            stats = compute_stats(result["steps"])
                            stats["total_error_ms"] = result["total_error_ms"]
                            stats["lag_stats"] = result["lag_stats"]
                            stats["action_type"] = result.get("action_type", action_type)
                            run_stats_list.append(stats)

                        print()  # 换行
                        all_results[(preset_name, action_type, delay_ms, steps)] = run_stats_list

    except KeyboardInterrupt:
        print("\n[中断] 用户取消测试")
    finally:
        if SIMULATE_CPU_LOAD:
            stop_cpu_load()

    # ── 输出汇总报告 ──
    print_summary_report(all_results)

    # ── 导出 CSV ──
    if EXPORT_CSV:
        export_csv_results(all_results)


def print_summary_report(all_results):
    """输出汇总对比报告"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(OUTPUT_DIR, f"report_{timestamp}.txt")

    lines = []
    def p(text=""):
        print(text)
        lines.append(text)

    p("\n" + "=" * 100)
    p("  精度对比汇总报告")
    p("=" * 100)

    # ── 按延迟分组对比 ──
    for action_type in TEST_ACTION_TYPES:
        for delay_ms in TEST_DELAY_MS:
            for steps in TEST_STEPS_PER_RUN:
                p(f"\n{'━' * 100}")
                p(f"  输入类型 = {action_type} | 延迟 = {delay_ms}ms | 步数 = {steps} | 重复 = {TEST_REPEATS}")
                p(f"{'━' * 100}")

                header = (f"{'配置方案':<20} │ "
                          f"{'平均误差(ms)':<14} │ {'中位误差(ms)':<14} │ "
                          f"{'σ误差(ms)':<12} │ {'P99误差(ms)':<12} │ "
                          f"{'P95误差(ms)':<12} │ {'最终偏差(ms)':<14} │ "
                          f"{'最大偏差(ms)':<14} │ {'总误差(ms)':<12}")
                p(header)
                p("─" * len(header))

                for preset_name in CONFIG_PRESETS:
                    key = (preset_name, action_type, delay_ms, steps)
                    runs = all_results.get(key, [])
                    if not runs:
                        p(f"  {preset_name:<20} │ (无数据)")
                        continue

                    # 跨 run 聚合
                    all_mean_errors = [r["mean_error_ms"] for r in runs]
                    all_median_errors = [r["median_error_ms"] for r in runs]
                    all_std_errors = [r["std_error_ms"] for r in runs]
                    all_p99 = [r["p99_error_ms"] for r in runs]
                    all_p95 = [r["p95_error_ms"] for r in runs]
                    all_final_bias = [r["final_cum_bias_ms"] for r in runs]
                    all_max_bias = [r["max_cum_bias_ms"] for r in runs]
                    all_total_err = [r["total_error_ms"] for r in runs]

                    p(f"  {preset_name:<20} │ "
                      f"{statistics.mean(all_mean_errors):>+10.4f}    │ "
                      f"{statistics.mean(all_median_errors):>+10.4f}    │ "
                      f"{statistics.mean(all_std_errors):>10.4f}  │ "
                      f"{statistics.mean(all_p99):>10.4f}  │ "
                      f"{statistics.mean(all_p95):>10.4f}  │ "
                      f"{statistics.mean(all_final_bias):>+10.4f}    │ "
                      f"{statistics.mean(all_max_bias):>10.4f}    │ "
                      f"{statistics.mean(all_total_err):>+10.4f}  ")

    # ── 按重复次数折算稳定性: 总量/TEST_REPEATS 与 每项均值 ──
    p(f"\n{'═' * 100}")
    p(f"  重复稳定性检查 (总量 ÷ 重复次数={TEST_REPEATS})")
    p(f"{'═' * 100}")

    summary_metric_defs = [
        ("mean_error_ms", "平均误差"),
        ("median_error_ms", "中位误差"),
        ("std_error_ms", "σ误差"),
        ("p95_error_ms", "P95误差"),
        ("p99_error_ms", "P99误差"),
        ("final_cum_bias_ms", "最终偏差"),
        ("max_cum_bias_ms", "最大偏差"),
        ("total_error_ms", "总误差"),
    ]

    for action_type in TEST_ACTION_TYPES:
        p(f"\n{'─' * 100}")
        p(f"  输入类型 = {action_type}")
        p("  说明: 折算总量 = 所有运行该指标求和 / TEST_REPEATS")
        p("       每项均值 = 折算总量 / (延迟梯度数 × 步数梯度数)")
        p(f"{'─' * 100}")

        for preset_name in CONFIG_PRESETS:
            preset_runs = []
            for (pn, at, dm, st), runs in all_results.items():
                if pn == preset_name and at == action_type and runs:
                    preset_runs.extend(runs)

            case_count = len(TEST_DELAY_MS) * len(TEST_STEPS_PER_RUN)
            if not preset_runs:
                p(f"  [{preset_name}] (无数据)")
                continue

            p(f"  [{preset_name}]")
            for metric_key, metric_label in summary_metric_defs:
                metric_values = [r[metric_key] for r in preset_runs if metric_key in r]
                if not metric_values:
                    continue
                repeat_normalized_total = sum(metric_values) / max(1, TEST_REPEATS)
                per_case_mean = repeat_normalized_total / max(1, case_count)
                p(f"    {metric_label:<10} -> 折算总量: {repeat_normalized_total:+.4f} ms | 每项均值: {per_case_mean:+.4f} ms")

    # ── 跨延迟综合排名 ──
    p(f"\n{'═' * 100}")
    p("  综合排名 (按平均绝对误差，跨所有延迟/步数)")
    p(f"{'═' * 100}")

    preset_scores = {}
    for preset_name in CONFIG_PRESETS:
        all_abs_errors = []
        for (pn, at, dm, st), runs in all_results.items():
            if pn != preset_name:
                continue
            for r in runs:
                all_abs_errors.append(abs(r["mean_error_ms"]))
                all_abs_errors.append(abs(r["final_cum_bias_ms"]))
        if all_abs_errors:
            preset_scores[preset_name] = statistics.mean(all_abs_errors)

    ranked = sorted(preset_scores.items(), key=lambda x: x[1])
    p(f"\n  {'排名':<6}{'配置方案':<25}{'平均绝对误差(ms)':<20}{'评价'}")
    p("  " + "─" * 60)
    for i, (name, score) in enumerate(ranked, 1):
        if i == 1:
            tag = "🏆 最优"
        elif score < ranked[0][1] * 1.2:
            tag = "✅ 优秀"
        elif score < ranked[0][1] * 1.5:
            tag = "⚠️ 一般"
        else:
            tag = "❌ 较差"
        p(f"  {i:<6}{name:<25}{score:<20.4f}{tag}")

    # ── 各配置单项最优 ──
    p(f"\n{'═' * 100}")
    p("  各维度最优配置")
    p(f"{'═' * 100}")

    dimensions = {
        "最低平均误差": lambda r: statistics.mean([x["mean_error_ms"] for x in r]),
        "最低累计偏差": lambda r: statistics.mean([abs(x["final_cum_bias_ms"]) for x in r]),
        "最低抖动(σ)": lambda r: statistics.mean([x["std_error_ms"] for x in r]),
        "最低P99尾部": lambda r: statistics.mean([x["p99_error_ms"] for x in r]),
        "最低总误差": lambda r: statistics.mean([abs(x["total_error_ms"]) for x in r]),
    }

    for dim_name, dim_fn in dimensions.items():
        best_preset = None
        best_val = float("inf")
        for preset_name in CONFIG_PRESETS:
            all_runs = []
            for (pn, at, dm, st), runs in all_results.items():
                if pn == preset_name and runs:
                    all_runs.extend(runs)
            if not all_runs:
                continue
            val = dim_fn(all_runs)
            if val < best_val:
                best_val = val
                best_preset = preset_name
        if best_preset:
            p(f"  {dim_name:<16} → {best_preset:<25} ({best_val:+.4f}ms)")

    p("\n" + "=" * 100)

    # 写入文件
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"\n[输出] 报告已保存: {report_path}")


def export_csv_results(all_results):
    """导出 CSV 详细数据"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(OUTPUT_DIR, f"detail_{timestamp}.csv")

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "配置方案", "输入类型", "延迟(ms)", "步数", "运行序号",
            "平均误差(ms)", "中位误差(ms)", "σ误差(ms)",
            "最大误差(ms)", "最小误差(ms)", "P95误差(ms)", "P99误差(ms)",
            "最终累计偏差(ms)", "最大绝对偏差(ms)", "总误差(ms)",
            "PI补偿次数", "PI总补偿量(ms)",
        ])
        for (preset_name, action_type, delay_ms, steps), runs in all_results.items():
            for run_idx, r in enumerate(runs):
                lag = r.get("lag_stats")
                writer.writerow([
                    preset_name, action_type, delay_ms, steps, run_idx + 1,
                    f"{r['mean_error_ms']:.4f}",
                    f"{r['median_error_ms']:.4f}",
                    f"{r['std_error_ms']:.4f}",
                    f"{r['max_error_ms']:.4f}",
                    f"{r['min_error_ms']:.4f}",
                    f"{r['p95_error_ms']:.4f}",
                    f"{r['p99_error_ms']:.4f}",
                    f"{r['final_cum_bias_ms']:.4f}",
                    f"{r['max_cum_bias_ms']:.4f}",
                    f"{r['total_error_ms']:.4f}",
                    lag["comp_count"] if lag else 0,
                    f"{lag['compensated_ms']:.4f}" if lag else 0,
                ])

    print(f"[输出] CSV 已保存: {csv_path}")


# ==============================================================================
# 入口
# ==============================================================================

if __name__ == "__main__":
    # 高级用法: 命令行参数
    # python macro_precision_benchmark.py --load   (开启CPU负载)
    # python macro_precision_benchmark.py --quick  (快速模式: 少量测试)
    if "--load" in sys.argv:
        SIMULATE_CPU_LOAD = True
    if "--quick" in sys.argv:
        TEST_REPEATS = 1
        TEST_DELAY_MS = [5, 50]
        TEST_STEPS_PER_RUN = [50]

    run_benchmark()

