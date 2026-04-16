#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
宏延迟调试脚本

测试项目：
1. 宏加载延迟（load_macro_from_file）
2. 宏编译延迟（compile_macro）
3. 等待延迟与波动（precise_sleep）
4. 按键输入调用延迟（send_key_event down/up）
5. 鼠标输入调用延迟（send_mouse_click md/mu）

说明：
- 按键输入项测的是“注入函数调用耗时”，不是游戏内实际生效时间。
- wait 项测的是“计划睡眠时长 vs 实际阻塞时长”。
"""

import argparse
import importlib.util
import math
import os
import statistics
import threading
import time
from typing import Dict, Iterable, List, Optional, Tuple


def percentile(values: List[float], q: float) -> float:
    if not values:
        return float("nan")
    if len(values) == 1:
        return values[0]
    arr = sorted(values)
    pos = (len(arr) - 1) * q
    lo = int(math.floor(pos))
    hi = int(math.ceil(pos))
    if lo == hi:
        return arr[lo]
    frac = pos - lo
    return arr[lo] + (arr[hi] - arr[lo]) * frac


def stats(values_ms: List[float]) -> Dict[str, float]:
    if not values_ms:
        return {
            "count": 0,
            "min": float("nan"),
            "max": float("nan"),
            "mean": float("nan"),
            "stdev": float("nan"),
            "p50": float("nan"),
            "p95": float("nan"),
            "p99": float("nan"),
            "range": float("nan"),
        }

    mean_v = statistics.fmean(values_ms)
    return {
        "count": len(values_ms),
        "min": min(values_ms),
        "max": max(values_ms),
        "mean": mean_v,
        "stdev": statistics.stdev(values_ms) if len(values_ms) > 1 else 0.0,
        "p50": percentile(values_ms, 0.50),
        "p95": percentile(values_ms, 0.95),
        "p99": percentile(values_ms, 0.99),
        "range": max(values_ms) - min(values_ms),
    }


def fmt_ms(v: float) -> str:
    if math.isnan(v):
        return "nan"
    return f"{v:.4f}"


def print_stats_block(title: str, data: Dict[str, float], ideal_desc: str, extra: str = "") -> None:
    print(f"\n[{title}]")
    print("样本={count} | 理想值: {ideal_desc}".format(count=int(data["count"]), ideal_desc=ideal_desc))
    if extra:
        print(extra)
    print(
        "min={min}ms | max={max}ms | mean={mean}ms | std={stdev}ms | "
        "p50={p50}ms | p95={p95}ms | p99={p99}ms | 波动范围={range}ms".format(
            min=fmt_ms(data["min"]),
            max=fmt_ms(data["max"]),
            mean=fmt_ms(data["mean"]),
            stdev=fmt_ms(data["stdev"]),
            p50=fmt_ms(data["p50"]),
            p95=fmt_ms(data["p95"]),
            p99=fmt_ms(data["p99"]),
            range=fmt_ms(data["range"]),
        )
    )


def load_framework_module(framework_path: str):
    spec = importlib.util.spec_from_file_location("macro_framework_debug", framework_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载框架文件: {framework_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def measure_macro_load(framework, macro_path: str, rounds: int) -> Tuple[List[float], Optional[List]]:
    samples = []
    latest_script = None
    for _ in range(rounds):
        t0 = time.perf_counter()
        latest_script = framework.load_macro_from_file(macro_path)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0
        samples.append(elapsed_ms)
    return samples, latest_script


def measure_macro_compile(framework, script: List, rounds: int) -> Tuple[List[float], Optional[List]]:
    samples = []
    latest_compiled = None
    for _ in range(rounds):
        t0 = time.perf_counter()
        latest_compiled = framework.compile_macro(script)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0
        samples.append(elapsed_ms)
    return samples, latest_compiled


def measure_wait_delays(framework, wait_targets_ms: Iterable[float], rounds_each: int) -> Dict[float, Dict[str, List[float]]]:
    result = {}
    for target_ms in wait_targets_ms:
        actual_samples = []
        error_samples = []
        for _ in range(rounds_each):
            stop_event = threading.Event()
            t0 = time.perf_counter()
            framework.precise_sleep(max(0.0, target_ms) / 1000.0, stop_event)
            actual_ms = (time.perf_counter() - t0) * 1000.0
            actual_samples.append(actual_ms)
            error_samples.append(actual_ms - target_ms)
        result[target_ms] = {
            "actual_ms": actual_samples,
            "error_ms": error_samples,
        }
    return result


def measure_key_latency(framework, vk_code: int, rounds: int, inner_repeats: int = 1) -> Dict[str, List[float]]:
    down_ms = []
    up_ms = []
    pair_ms = []
    down_cpu_ms = []
    up_cpu_ms = []
    pair_cpu_ms = []

    thread_time_ns = getattr(time, "thread_time_ns", None)

    def _cpu_ms() -> float:
        if thread_time_ns is not None:
            return thread_time_ns() / 1_000_000.0
        return time.thread_time() * 1000.0

    repeats = max(1, inner_repeats)
    for _ in range(rounds):
        t0 = time.perf_counter()
        c0 = _cpu_ms()
        for _i in range(repeats):
            framework.send_key_event(vk_code, False)
        t1 = time.perf_counter()
        c1 = _cpu_ms()
        for _i in range(repeats):
            framework.send_key_event(vk_code, True)
        t2 = time.perf_counter()
        c2 = _cpu_ms()

        down_ms.append(((t1 - t0) * 1000.0) / repeats)
        up_ms.append(((t2 - t1) * 1000.0) / repeats)
        pair_ms.append(((t2 - t0) * 1000.0) / repeats)
        down_cpu_ms.append((c1 - c0) / repeats)
        up_cpu_ms.append((c2 - c1) / repeats)
        pair_cpu_ms.append((c2 - c0) / repeats)

    return {
        "down_ms": down_ms,
        "up_ms": up_ms,
        "pair_ms": pair_ms,
        "down_cpu_ms": down_cpu_ms,
        "up_cpu_ms": up_cpu_ms,
        "pair_cpu_ms": pair_cpu_ms,
    }


def measure_mouse_latency(framework, left: bool, rounds: int, inner_repeats: int = 1) -> Dict[str, List[float]]:
    down_ms = []
    up_ms = []
    pair_ms = []
    down_cpu_ms = []
    up_cpu_ms = []
    pair_cpu_ms = []

    thread_time_ns = getattr(time, "thread_time_ns", None)

    def _cpu_ms() -> float:
        if thread_time_ns is not None:
            return thread_time_ns() / 1_000_000.0
        return time.thread_time() * 1000.0

    repeats = max(1, inner_repeats)
    for _ in range(rounds):
        t0 = time.perf_counter()
        c0 = _cpu_ms()
        for _i in range(repeats):
            framework.send_mouse_click(left=left, up=False)  # md
        t1 = time.perf_counter()
        c1 = _cpu_ms()
        for _i in range(repeats):
            framework.send_mouse_click(left=left, up=True)   # mu
        t2 = time.perf_counter()
        c2 = _cpu_ms()

        down_ms.append(((t1 - t0) * 1000.0) / repeats)
        up_ms.append(((t2 - t1) * 1000.0) / repeats)
        pair_ms.append(((t2 - t0) * 1000.0) / repeats)
        down_cpu_ms.append((c1 - c0) / repeats)
        up_cpu_ms.append((c2 - c1) / repeats)
        pair_cpu_ms.append((c2 - c0) / repeats)

    return {
        "down_ms": down_ms,
        "up_ms": up_ms,
        "pair_ms": pair_ms,
        "down_cpu_ms": down_cpu_ms,
        "up_cpu_ms": up_cpu_ms,
        "pair_cpu_ms": pair_cpu_ms,
    }


def tune_runtime_for_measurement(framework, boost_priority: bool, affinity_core: Optional[int]) -> List[str]:
    notes: List[str] = []
    if boost_priority and hasattr(framework, "boost_thread_priority"):
        try:
            framework.boost_thread_priority()
            notes.append("线程优先级=TIME_CRITICAL")
        except Exception as e:
            notes.append(f"线程优先级提升失败: {e}")

    if affinity_core is not None and hasattr(framework, "set_thread_affinity"):
        try:
            framework.set_thread_affinity(affinity_core)
            notes.append(f"线程亲和性=CPU{affinity_core}")
        except Exception as e:
            notes.append(f"线程亲和性设置失败: {e}")

    return notes


def set_input_backend(framework, backend: str) -> bool:
    """切换输入后端。backend: interception|sendinput"""
    if backend == "sendinput":
        framework.USE_INTERCEPTION = False
        return True

    if backend == "interception":
        framework.USE_INTERCEPTION = True
        framework._interception_available = None
        framework._interception_send_key = None
        framework._interception_send_mouse = None
        if not hasattr(framework, "init_interception_backend"):
            return False
        try:
            return bool(framework.init_interception_backend())
        except Exception:
            return False

    return False


def detect_available_backends(framework) -> List[str]:
    backends = ["sendinput"]
    if set_input_backend(framework, "interception"):
        backends.insert(0, "interception")
    return backends


def parse_wait_points(text: str) -> List[float]:
    values = []
    for part in text.split(","):
        p = part.strip()
        if not p:
            continue
        values.append(float(p))
    return values


def get_input_latency_ideal(backend: str) -> Dict[str, str]:
    """按输入后端返回建议阈值文案。"""
    if backend == "interception":
        return {
            "single": "调用建议 mean <= 0.03ms",
            "pair": "调用建议 mean <= 0.06ms",
        }
    return {
        "single": "调用建议 mean <= 0.3ms",
        "pair": "调用建议 mean <= 0.6ms",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="宏延迟调试脚本：测量加载/等待/键鼠注入延迟与波动")
    parser.add_argument("--macro", default=os.path.join("宏", "测试.json"), help="待测试宏文件路径")
    parser.add_argument("--load-rounds", type=int, default=30, help="宏加载测试轮数")
    parser.add_argument("--compile-rounds", type=int, default=30, help="宏编译测试轮数")
    parser.add_argument("--wait-rounds", type=int, default=100, help="每个wait目标时长的测试轮数")
    parser.add_argument("--wait-points", default="1,2,5,10,20,50", help="wait目标毫秒，逗号分隔")
    parser.add_argument("--key-rounds", type=int, default=200, help="按键注入测试轮数")
    parser.add_argument("--mouse-rounds", type=int, default=200, help="鼠标注入测试轮数")
    parser.add_argument("--mouse-button", choices=["left", "right"], default="left", help="鼠标测试按键")
    parser.add_argument("--vk", type=lambda x: int(x, 0), default=0xFC, help="测试用VK码，默认0xFC(未绑定的OEM键)")
    parser.add_argument("--foreground-friendly", action="store_true", help="前台干预友好模式：输出CPU时延并进行运行时调优")
    parser.add_argument("--inner-repeats", type=int, default=1, help="单样本内部重复发送次数（>1可显著降低前台干预抖动）")
    parser.add_argument("--affinity-core", type=int, default=None, help="将测量线程绑定到指定CPU核心（如 3）")
    parser.add_argument("--no-boost-priority", action="store_true", help="不提升测量线程优先级")
    args = parser.parse_args()

    project_dir = os.path.dirname(os.path.abspath(__file__))
    framework_path = os.path.join(project_dir, "宏执行框架1.9.py")
    macro_path = args.macro
    if not os.path.isabs(macro_path):
        macro_path = os.path.join(project_dir, macro_path)

    if not os.path.exists(framework_path):
        raise FileNotFoundError(f"未找到框架文件: {framework_path}")
    if not os.path.exists(macro_path):
        raise FileNotFoundError(f"未找到宏文件: {macro_path}")

    framework = load_framework_module(framework_path)

    original_use_interception = getattr(framework, "USE_INTERCEPTION", False)

    print("=" * 72)
    print("宏延迟调试开始")
    print(f"框架: {framework_path}")
    print(f"宏文件: {macro_path}")
    print(f"USE_INTERCEPTION: {getattr(framework, 'USE_INTERCEPTION', 'unknown')}")
    print("=" * 72)
    if args.foreground_friendly:
        print("模式: 前台干预友好模式（允许前台微调，建议关注CPU时延统计）")
        if args.inner_repeats <= 1:
            print("建议: 前台干预时可加 --inner-repeats 20 以降低抖动")
    else:
        print("模式: 严格基准模式（建议全程不操作键鼠，获得最低波动）")
    print("提示: 输入注入测试会发往前台窗口；建议切到目标窗口后再测试，避免终端蜂鸣或误触。")

    available_backends = detect_available_backends(framework)
    print(
        "[初始化] 可用输入后端: "
        + ", ".join("Interception" if x == "interception" else "SendInput" for x in available_backends)
    )

    load_samples, script = measure_macro_load(framework, macro_path, max(1, args.load_rounds))
    if script is None:
        raise RuntimeError("宏加载失败，无法继续")

    compile_samples, compiled = measure_macro_compile(framework, script, max(1, args.compile_rounds))
    if compiled is None:
        raise RuntimeError("宏编译失败，无法继续")

    wait_points = parse_wait_points(args.wait_points)
    wait_data = measure_wait_delays(framework, wait_points, max(1, args.wait_rounds))

    tune_notes = tune_runtime_for_measurement(
        framework,
        boost_priority=not args.no_boost_priority,
        affinity_core=args.affinity_core,
    )

    input_results = {}
    mouse_left = args.mouse_button == "left"
    for backend in available_backends:
        ok = set_input_backend(framework, backend)
        if not ok:
            print(f"[输入测试] 跳过后端: {backend}（初始化失败）")
            continue
        key_data = measure_key_latency(
            framework,
            args.vk,
            max(1, args.key_rounds),
            inner_repeats=max(1, args.inner_repeats),
        )
        mouse_data = measure_mouse_latency(
            framework,
            mouse_left,
            max(1, args.mouse_rounds),
            inner_repeats=max(1, args.inner_repeats),
        )
        input_results[backend] = {
            "key": key_data,
            "mouse": mouse_data,
        }

    # 还原默认配置，避免影响其他脚本运行。
    framework.USE_INTERCEPTION = original_use_interception

    raw_count = len(script)
    compiled_count = len(compiled)

    load_stats = stats(load_samples)
    compile_stats = stats(compile_samples)

    print("\n" + "-" * 72)
    print("总览")
    print(f"宏原始指令数: {raw_count}")
    print(f"宏编译后动作数: {compiled_count}")
    print("-" * 72)
    if tune_notes:
        print("测量线程调优: " + " | ".join(tune_notes))

    load_norm = (load_stats["mean"] / max(1, raw_count)) * 100.0
    print_stats_block(
        "宏加载延迟",
        load_stats,
        "普通本地JSON建议 mean <= 5ms（小宏）",
        extra=f"折算每100条原始指令平均加载耗时: {load_norm:.4f}ms",
    )

    compile_norm = (compile_stats["mean"] / max(1, compiled_count)) * 100.0
    print_stats_block(
        "宏编译延迟",
        compile_stats,
        "建议 mean <= 3ms（小到中等宏）",
        extra=f"折算每100条编译后动作平均编译耗时: {compile_norm:.4f}ms",
    )

    print("\n[等待延迟(wait/precise_sleep)]")
    print("理想值: 实测值尽量接近目标值，误差均值接近0ms，std/p95越小越稳定")
    for target in wait_points:
        samples = wait_data[target]
        actual_stats = stats(samples["actual_ms"])
        error_stats = stats(samples["error_ms"])
        print(
            "目标={target:.3f}ms | 实测mean={amean}ms p95={ap95}ms std={astd}ms | "
            "误差mean={emean}ms p95={ep95}ms p99={ep99}ms 范围={erange}ms".format(
                target=target,
                amean=fmt_ms(actual_stats["mean"]),
                ap95=fmt_ms(actual_stats["p95"]),
                astd=fmt_ms(actual_stats["stdev"]),
                emean=fmt_ms(error_stats["mean"]),
                ep95=fmt_ms(error_stats["p95"]),
                ep99=fmt_ms(error_stats["p99"]),
                erange=fmt_ms(error_stats["range"]),
            )
        )

    for backend, payload in input_results.items():
        name = "Interception" if backend == "interception" else "SendInput"
        key_data = payload["key"]
        mouse_data = payload["mouse"]
        ideal = get_input_latency_ideal(backend)

        print(f"\n[输入延迟后端: {name}]")

        print_stats_block(
            "按键输入延迟(key down)",
            stats(key_data["down_ms"]),
            ideal["single"],
        )
        print_stats_block(
            "按键输入延迟(key up)",
            stats(key_data["up_ms"]),
            ideal["single"],
        )
        print_stats_block(
            "按键输入往返延迟(down+up)",
            stats(key_data["pair_ms"]),
            ideal["pair"],
        )
        if args.foreground_friendly:
            print_stats_block(
                "按键CPU时延(down+up)",
                stats(key_data["pair_cpu_ms"]),
                "该值更接近注入函数自身开销，受前台抢占影响更小",
            )

        # 鼠标部分
        print_stats_block(
            "鼠标输入延迟(md)",
            stats(mouse_data["down_ms"]),
            ideal["single"],
            extra=f"测试按键: {args.mouse_button}",
        )
        print_stats_block(
            "鼠标输入延迟(mu)",
            stats(mouse_data["up_ms"]),
            ideal["single"],
            extra=f"测试按键: {args.mouse_button}",
        )
        print_stats_block(
            "鼠标输入往返延迟(md+mu)",
            stats(mouse_data["pair_ms"]),
            ideal["pair"],
            extra=f"测试按键: {args.mouse_button}",
        )
        if args.foreground_friendly:
            print_stats_block(
                "鼠标CPU时延(md+mu)",
                stats(mouse_data["pair_cpu_ms"]),
                "该值更接近注入函数自身开销，受前台抢占影响更小",
                extra=f"测试按键: {args.mouse_button}",
            )
        

    print("\n" + "=" * 72)
    print("调试完成")
    print("说明: 若需要逼近理想值，可提高进程/线程优先级，减少后台负载，并固定CPU亲和性后复测。")
    print("=" * 72)


if __name__ == "__main__":
    main()
