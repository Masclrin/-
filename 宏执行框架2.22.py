# -*- coding: utf-8 -*-
"""
[通用宏执行框架 - 鼠标侧键+键盘双触发版]
核心特性：
1. 兼容原有JSON宏格式（kd/ku/md/mu/view/loop/wait/嵌套import）
2. 执行宏过程中允许物理鼠标/键盘干预操作
3. 键盘按键 + 鼠标侧键 绑定宏文件，按下即执行
4. 高精度输入模拟 + 非阻塞式执行，平衡精度与交互性
5. ESC一键终止当前宏，防止卡键
"""
import os
import re
import time
import threading
import ctypes
import json
import random
import gc
from collections import deque
from precision_engine_v5 import (
    boost_process_priority,
    boost_thread_priority,
    set_thread_affinity,
    precise_sleep_v5 as precise_sleep,
    auto_warmup,
)
from pynput import keyboard, mouse

# —— 配置自定义常量 ——
# 宏文件绑定 (键盘名 / 鼠标侧键（mouse_x1/mouse_x2）: 宏文件路径)
MACRO_BINDINGS = {
    "caps_lock": "宏/按键测试.json",
    #"v": "宏/测试.json",
    #"f2": "宏/宏2.json",
    #"mouse_x1": "宏/ZQzsZzz.json",   # 鼠标后退侧键(4号)
    #"mouse_x2": "宏/侧键宏2.json",    # 鼠标前进侧键(5号)
}

# 每个触发键可单独配置执行方式（可只配置需要的键）
# repeat: 执行次数（缺省/非法/<=1 时默认单次）
# running_press_mode（宏按键触发方式）:
#   normal -> 常规按下执行  |  pause_resume -> 切换暂停/继续  |  hold_pause -> 按住暂停，抬起继续
#   stop_restart -> 重新启动实例  |  parallel_trigger -> 多次按下并发新增实例
#   release_stop -> 抬起终止，按下重开  |  release_pause -> 抬起暂停，按下继续
MACRO_RUNTIME_SETTINGS = {
    "caps_lock": {"repeat": 1, "running_press_mode": "normal"},
    # "v": {"repeat": 3, "running_press_mode": "normal"},
    # "mouse_x1": {"repeat": 999999, "running_press_mode": "release_pause"},
}

MACRO_RELOAD_KEY = None              # 重新加载宏配置 keyboard.Key.f6 或 "f6";可None
STOP_MACRO_KEY = keyboard.Key.esc    # 宏终止按键
GLOBAL_PAUSE_TOGGLE_KEY = "f8"       # 全局暂停开关键
ENABLE_TEST_TRACE = True             # 测试功能开关：打印每步动作与耗时明细
TEST_TRACE_OUTPUT_MODE = "final"     # 测试输出模式: realtime(逐行实时，不建议。精度波动严重) / final(执行后统一输出)

# 精度优化
USE_INTERCEPTION = True              # Interception总开关
USE_INTERCEPTION_KEYBOARD = True     # 键盘是否使用 Interception
USE_INTERCEPTION_MOUSE = True        # 鼠标是否使用 Interception（左键吞输入时建议 False）
LAG_COMPENSATION_ENABLED = True      # 累计误差补偿总开关，默认使用前馈补偿

# 调度优化开关
MAX_MACRO_EXEC_TIME = 25000          # 单个宏最大执行时间（秒），防止死循环
MAX_COMPILED_ACTIONS = 200000        # 单次编译允许的最大动作数（防止超大loop展开卡死）

# —— 高级参数 ——
ENABLE_REALTIME_PRIORITY = True      # 进程优先级切换到 REALTIME（谨慎开启）
ENABLE_PRO_AUDIO = True              # 工作线程注入 MMCSS Pro Audio

ENABLE_RANDOM_DELAY_ADJUST = False   # 随机延迟开关
# 按该动作原始延迟（毫秒）分档，直接修正"动作原延迟本身"（可增可减）
# 1) 同时配置固定值与百分比时，最终修正 = 固定值修正 + 百分比修正：
#    - fixed_ms：固定值（可正可负）- percent：固定百分比（可正可负）
#    - fixed_ms_down + fixed_ms_up：在[-fixed_ms_down, +fixed_ms_up]内随机（down/up支持配置为负值）
#    - percent_down + percent_up：在[-percent_down, +percent_up]内随机（down/up支持配置为负值）
RANDOM_DELAY_ADJUST_RULES_BY_DELAY_MS = [
    {"min": 0, "max": 160, "fixed_ms_down": 5, "fixed_ms_up": 10},
    {"min": 160, "max": 490, "percent_down": 3, "percent_up": 5},
    {"min": 580, "max": 900, "fixed_ms": -50},
]

LAG_USE_PI_CONTROLLER = False         # 是否启用PI控制器进行误差补偿
LAG_USE_FEEDFORWARD = True            # 是否启用前馈控制进行补偿（通常只开一个）

# PI 控制器参数
LAG_FILTER_PROCESS_NOISE = 1e-2       # 延迟滤波过程噪声
LAG_FILTER_MEASURE_NOISE = 2e-2       # 延迟滤波测量噪声
LAG_ERROR_TRIGGER_MS = 0.1            # 累计误差>(ms) → 开始补偿
LAG_ERROR_TARGET_MS = 0.05            # 补偿目标，达到后停止
LAG_MAX_STEP_COMP_PCT = 0.02          # 单步补偿不超过该步延迟的 2%（防抖动）
LAG_MIN_STEP_DELAY_MS = 0.3           # 单步最低保留延迟（ms），不会压到 0
LAG_KP = 0.25                         # 比例增益：响应当前误差幅度（越大越激进）
LAG_KI = 0.15                         # 积分增益：持续累加误差直到补偿=开销（核心项）
LAG_INTEGRAL_MAX = 1.0                # 积分上限(秒·步)，防止极端情况积分爆炸
LAG_INTEGRAL_DECAY = 0.996            # 误差低于触发值时积分衰减系数（避免残留）

# 前馈+比例控制器参数
FF_WINDOW_SIZE = 16                  # 前馈滑动窗口大小
FF_WARMUP_STEPS = 2                  # 前馈启动所需最小样本数
FF_SCALE = 0.98                      # 前馈缩放因子（0.95~1.0，略低于1可防负偏）
FF_KP = 0.85                         # 比例增益（0.15~0.90）高值会偏负数
FF_DEAD_ZONE_MS = 0.005              # 死区：误差小于此值不修正（ms）
FF_MAX_COMP_PCT = 0.15               # 单步最大补偿占延迟比例（5%~15%）
FF_RATE_LIMIT_PCT = 0.20             # 修正值每步最大变化率
FF_ANTI_OVERSHOOT = 0.5              # 误差过零时衰减系数（0.05~0.30，越小越保守）
FF_EMA_ALPHA = 0.25                  # 误差平滑EMA系数（0.15~0.40）

# 全局状态（支持多宏并发）
active_macros = {}                    # macro_id -> {thread, stop_event, trigger_key, macro_name}
trigger_to_macro_ids = {}             # trigger_key -> set(macro_id)
starting_triggers = set()             # 正在加载/编译中的触发键，防止短时间重复触发
active_macros_lock = threading.Lock()
macro_cache_lock = threading.Lock()
compiled_macro_cache = {}             # trigger_key -> {compiled_macro, macro_name, macro_path, loaded_at}
macro_cache_ready = False             # 已完成一次全量预加载
macro_loading_in_progress = False    # True 时禁止触发新宏，避免加载过程中触发
pressed_trigger_keys = set()          # 防止键盘按住时重复触发
macro_id_seed = 0
global_trigger_paused = False         # True 时禁止所有宏触发（仅允许全局暂停键解锁）

VALID_RUNNING_PRESS_MODES = frozenset({
    "normal", "pause_resume", "hold_pause", "stop_restart",
    "parallel_trigger", "release_stop", "release_pause",
})

# ==============================================================================
# 底层WinAPI定义
# ==============================================================================
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
winmm = ctypes.windll.winmm

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010

# WinAPI 结构体定义
class _INPUT(ctypes.Structure):
    pass
_INPUT._fields_ = [
    ("type", ctypes.c_ulong),
    ("mi", (ctypes.c_long * 3)),
    ("mouseData", ctypes.c_ulong),
    ("dwFlags", ctypes.c_ulong),
    ("time", ctypes.c_ulong),
    ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
]

# 按键映射表（预计算VK值用于快速查找）
VK_MAP = {
    "left": 0x01, "right": 0x02, "cancel": 0x03, "middle": 0x04, "back": 0x08,
    "tab": 0x09, "enter": 0x0D, "shift": 0x10, "ctrl": 0x11, "alt": 0x12,
    "pause": 0x13, "caps_lock": 0x14, "esc": 0x1B, "space": 0x20, "page_up": 0x21,
    "page_down": 0x22, "end": 0x23, "home": 0x24, "left_arrow": 0x25,
    "up_arrow": 0x26, "right_arrow": 0x27, "down_arrow": 0x28, "print_screen": 0x2C,
    "insert": 0x2D, "delete": 0x2E,
    "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34, "5": 0x35,
    "6": 0x36, "7": 0x37, "8": 0x38, "9": 0x39,
    "a": 0x41, "b": 0x42, "c": 0x43, "d": 0x44, "e": 0x45, "f": 0x46,
    "g": 0x47, "h": 0x48, "i": 0x49, "j": 0x4A, "k": 0x4B, "l": 0x4C,
    "m": 0x4D, "n": 0x4E, "o": 0x4F, "p": 0x50, "q": 0x51, "r": 0x52,
    "s": 0x53, "t": 0x54, "u": 0x55, "v": 0x56, "w": 0x57, "x": 0x58,
    "y": 0x59, "z": 0x5A,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73,
    "f5": 0x74, "f6": 0x75, "f7": 0x76, "f8": 0x77,
    "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
}

# 预编译的正则表达式
_COMMENT_PATTERN = re.compile(r'//.*?$|/\*.*?\*/|\'(?:\\.|[^\\\'])*\'|"(?:\\.|[^\\"])*"', re.DOTALL | re.MULTILINE)


def get_vk(key):
    """获取按键对应的VK码"""
    return VK_MAP.get(str(key).lower(), 0)


# Interception 相关
_interception_available = None
_interception_send_key = None
_interception_send_mouse = None


def apply_realtime_priority(enable_realtime=True):
    """按开关设置进程优先级: True=REALTIME, False=保持当前。"""
    if not enable_realtime:
        return False
    try:
        pid = kernel32.GetCurrentProcessId()
        handle = kernel32.OpenProcess(0x1F0FFF, False, pid)
        if not handle:
            return False
        ok = bool(kernel32.SetPriorityClass(handle, 0x00000100))
        kernel32.CloseHandle(handle)
        return ok
    except Exception:
        return False


def apply_pro_audio(enable_pro_audio=True):
    """按开关注入 MMCSS Pro Audio 调度。"""
    if not enable_pro_audio:
        return False
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


def init_interception_backend():
    """初始化 Interception 后端（懒加载）"""
    global _interception_available, _interception_send_key, _interception_send_mouse

    if _interception_available is not None:
        return _interception_available

    if not USE_INTERCEPTION:
        _interception_available = False
        print("[系统] Interception 总开关关闭，键盘/鼠标均使用 SendInput")
        return False

    try:
        from interception_input import (
            get_interception_context,
            send_key_interception,
            send_mouse_interception,
        )
        get_interception_context()
        _interception_send_key = send_key_interception
        _interception_send_mouse = send_mouse_interception
        _interception_available = True
        print(f"[系统] Interception 驱动已预加载 | keyboard={'on' if USE_INTERCEPTION_KEYBOARD else 'off'}, "
              f"mouse={'on' if USE_INTERCEPTION_MOUSE else 'off'}")
    except Exception as e:
        _interception_available = False
        _interception_send_key = None
        _interception_send_mouse = None
        print(f"[系统] Interception 不可用，回退 SendInput，精度为1ms\n原因:{e}")

    return _interception_available


# 预创建的 SendInput 输入结构
class _SendInputStruct:
    __slots__ = ('ki', 'mi')

    def __init__(self):
        self.ki = (ctypes.c_ulong * 2)()
        self.ki[0] = 0  # wVk
        self.ki[1] = 0  # dwFlags
        self.mi = (ctypes.c_ulong * 7)()
        self.mi[0] = 0  # dx
        self.mi[1] = 0  # dy
        self.mi[2] = 0  # mouseData
        self.mi[3] = 0  # dwFlags
        self.mi[4] = 0  # time
        self.mi[5] = 0  # dwExtraInfo (placeholder)


def send_key_event(vk_code, up=False):
    """发送键盘事件（自动选择最高精度后端）"""
    if USE_INTERCEPTION and USE_INTERCEPTION_KEYBOARD:
        global _interception_available
        if _interception_available is None:
            init_interception_backend()

        if _interception_available and _interception_send_key:
            # VK → 扫描码 反查（缓存映射）
            key_name = _VK_TO_NAME.get(vk_code)
            if key_name and _interception_send_key(key_name, up):
                return

    # 回退到原有 SendInput
    inp = _SendInputStruct()
    inp.ki[0] = vk_code
    inp.ki[1] = KEYEVENTF_KEYUP if up else 0
    user32.SendInput(1, ctypes.byref(inp), 28)  # sizeof(INPUT) = 28


# 预计算 VK → 名称 映射
_VK_TO_NAME = {v: k for k, v in VK_MAP.items()}


def send_mouse_click(left=True, up=False):
    """发送鼠标点击事件"""
    if USE_INTERCEPTION and USE_INTERCEPTION_MOUSE:
        global _interception_available
        if _interception_available is None:
            init_interception_backend()

        if _interception_available and _interception_send_mouse:
            try:
                if _interception_send_mouse(left=left, up=up):
                    return
            except Exception as e:
                print(f"[系统] Interception 鼠标发送失败，回退 SendInput：{e}")

    inp = _SendInputStruct()
    if left:
        inp.mi[3] = MOUSEEVENTF_LEFTUP if up else MOUSEEVENTF_LEFTDOWN
    else:
        inp.mi[3] = MOUSEEVENTF_RIGHTUP if up else MOUSEEVENTF_RIGHTDOWN
    user32.SendInput(1, ctypes.byref(inp), 28)


def send_mouse_move(x, y):
    """发送鼠标移动事件（相对位移）"""
    inp = _SendInputStruct()
    inp.mi[0] = int(x)
    inp.mi[1] = int(y)
    inp.mi[3] = MOUSEEVENTF_MOVE
    user32.SendInput(1, ctypes.byref(inp), 28)


def set_cursor_pos(x, y):
    """设置鼠标绝对位置"""
    user32.SetCursorPos(int(x), int(y))


def next_macro_id():
    """生成递增宏实例ID（线程安全）"""
    global macro_id_seed
    with active_macros_lock:
        macro_id_seed += 1
        return macro_id_seed


def resolve_macro_config(trigger_key):
    """获取触发键对应的宏路径与运行配置（含默认值）"""
    if trigger_key not in MACRO_BINDINGS:
        return None

    runtime = MACRO_RUNTIME_SETTINGS.get(trigger_key, {})
    raw_repeat = runtime.get("repeat", 1)
    try:
        repeat_count = max(1, int(raw_repeat))
    except (TypeError, ValueError):
        repeat_count = 1

    running_press_mode = str(runtime.get("running_press_mode", "normal")).lower()
    if running_press_mode not in VALID_RUNNING_PRESS_MODES:
        running_press_mode = "normal"

    return {
        "macro_path": MACRO_BINDINGS[trigger_key],
        "repeat_count": repeat_count,
        "running_press_mode": running_press_mode,
    }


def register_active_macro(macro_id, trigger_key, macro_name, thread, stop_event, pause_event):
    with active_macros_lock:
        active_macros[macro_id] = {
            "thread": thread, "stop_event": stop_event, "pause_event": pause_event,
            "trigger_key": trigger_key, "macro_name": macro_name,
        }
        trigger_to_macro_ids.setdefault(trigger_key, set()).add(macro_id)


def _collect_live_macro_infos(trigger_key):
    """获取触发键对应的存活宏信息，并清理已死亡实例。"""
    with active_macros_lock:
        existing_ids = list(trigger_to_macro_ids.get(trigger_key, set()))
        live_infos, live_ids = [], []

        for mid in existing_ids:
            info = active_macros.get(mid)
            if not info:
                continue
            th = info.get("thread")
            if th and th.is_alive():
                live_infos.append(info)
                live_ids.append(mid)
            else:
                active_macros.pop(mid, None)

        if live_ids:
            trigger_to_macro_ids[trigger_key] = set(live_ids)
        else:
            trigger_to_macro_ids.pop(trigger_key, None)

        return live_infos


def release_starting_trigger(trigger_key):
    """释放触发键启动占位"""
    with active_macros_lock:
        starting_triggers.discard(trigger_key)


def try_reserve_trigger(trigger_key, allow_parallel=False):
    """为触发键申请启动占位。返回：(ok, reason)"""
    with active_macros_lock:
        if macro_loading_in_progress:
            return False, "宏配置加载中"
        if trigger_key in starting_triggers:
            return False, "宏正在启动中"

        existing_ids = list(trigger_to_macro_ids.get(trigger_key, set()))
        live_ids = []
        for mid in existing_ids:
            info = active_macros.get(mid)
            if not info:
                continue
            th = info.get("thread")
            if th and th.is_alive():
                live_ids.append(mid)
            else:
                active_macros.pop(mid, None)

        if live_ids and not allow_parallel:
            trigger_to_macro_ids[trigger_key] = set(live_ids)
            macro_name = active_macros[live_ids[0]].get("macro_name", "未知宏")
            return False, f"宏已在运行：{macro_name}"

        trigger_to_macro_ids.pop(trigger_key, None)
        starting_triggers.add(trigger_key)
        return True, ""


def unregister_active_macro(macro_id):
    with active_macros_lock:
        info = active_macros.pop(macro_id, None)
        if not info:
            return
        trigger_key = info.get("trigger_key")
        ids = trigger_to_macro_ids.get(trigger_key)
        if ids:
            ids.discard(macro_id)
            if not ids:
                trigger_to_macro_ids.pop(trigger_key, None)


def pause_macros_by_trigger(trigger_key, reason="按键触发暂停"):
    """暂停某个触发键启动的所有存活宏。"""
    with active_macros_lock:
        macro_ids = list(trigger_to_macro_ids.get(trigger_key, set()))
        pause_events = [active_macros[mid].get("pause_event") for mid in macro_ids if mid in active_macros]

    changed = sum(1 for ev in pause_events if ev is not None and not ev.is_set())
    for ev in pause_events:
        if ev is not None and not ev.is_set():
            ev.set()

    if changed:
        print(f"[宏暂停] {reason}：{trigger_key} -> 暂停{changed}个宏实例")
    return changed


def resume_macros_by_trigger(trigger_key, reason="按键触发继续"):
    """继续某个触发键启动的所有暂停宏。"""
    with active_macros_lock:
        macro_ids = list(trigger_to_macro_ids.get(trigger_key, set()))
        pause_events = [active_macros[mid].get("pause_event") for mid in macro_ids if mid in active_macros]

    changed = sum(1 for ev in pause_events if ev is not None and ev.is_set())
    for ev in pause_events:
        if ev is not None and ev.is_set():
            ev.clear()

    if changed:
        print(f"[宏继续] {reason}：{trigger_key} -> 继续{changed}个宏实例")
    return changed


def has_live_macro_for_trigger(trigger_key):
    """判断触发键是否已有存活宏实例。"""
    return bool(_collect_live_macro_infos(trigger_key))


def has_paused_macro_for_trigger(trigger_key):
    """判断触发键是否存在暂停态宏实例。"""
    return any(info.get("pause_event") is not None and info["pause_event"].is_set()
               for info in _collect_live_macro_infos(trigger_key))


def normalize_hotkey_to_name(hotkey):
    """将热键配置统一为字符串名称，失败返回None。"""
    if hotkey is None:
        return None
    if isinstance(hotkey, keyboard.Key):
        return hotkey.name
    if isinstance(hotkey, str):
        return hotkey.strip().lower() or None
    return None


def get_pressed_key_name(key):
    """将监听到的按键对象统一为字符串名称，失败返回None。"""
    if isinstance(key, keyboard.Key):
        return key.name
    if isinstance(key, keyboard.KeyCode):
        return key.char.lower() if key.char else None
    return None


def is_hotkey_match(configured_hotkey, pressed_key_name):
    """热键匹配：支持 keyboard.Key 与字符串配置。"""
    cfg_name = normalize_hotkey_to_name(configured_hotkey)
    return bool(cfg_name) and pressed_key_name == cfg_name


def is_reload_hotkey(key_str):
    reload_key = normalize_hotkey_to_name(MACRO_RELOAD_KEY)
    return bool(reload_key) and key_str == reload_key


def _set_macro_loading_state(loading):
    """统一更新宏加载状态。"""
    global macro_loading_in_progress
    with macro_cache_lock:
        macro_loading_in_progress = bool(loading)


def _build_compiled_macro_cache_for_all_bindings(reason="启动预加载"):
    """全量加载并编译绑定宏，写入缓存。"""
    global macro_cache_ready
    _set_macro_loading_state(True)
    start_ts = time.perf_counter()
    new_cache = {}
    ok_count = fail_count = 0

    try:
        print(f"[宏预编译] 开始{reason}，共{len(MACRO_BINDINGS)}个绑定")
        for trigger_key, macro_path in MACRO_BINDINGS.items():
            try:
                script = load_macro_from_file(macro_path)
                compiled = compile_macro(script)
                new_cache[trigger_key] = {
                    "compiled_macro": compiled,
                    "macro_name": os.path.basename(macro_path),
                    "macro_path": macro_path,
                    "loaded_at": time.strftime("%H:%M:%S"),
                }
                ok_count += 1
                print(f"[宏预编译] 成功：{trigger_key} -> {macro_path} | 指令数={len(compiled)}")
            except Exception as e:
                fail_count += 1
                print(f"[宏预编译] 失败：{trigger_key} -> {macro_path} | {e}")

        with macro_cache_lock:
            compiled_macro_cache.clear()
            compiled_macro_cache.update(new_cache)
            macro_cache_ready = True

        elapsed_ms = (time.perf_counter() - start_ts) * 1000.0
        print(f"[宏预编译] 完成：成功{ok_count}，失败{fail_count}，耗时{elapsed_ms:.1f}ms")
        return ok_count, fail_count
    finally:
        _set_macro_loading_state(False)


def preload_macros_if_enabled():
    """仅在配置了重载键时，启动阶段进行一次全量预加载。"""
    if normalize_hotkey_to_name(MACRO_RELOAD_KEY):
        _build_compiled_macro_cache_for_all_bindings("启动阶段预加载")
    else:
        print("[宏预编译] 未配置重载键，保持按触发键时即时加载/编译")


def reload_macro_cache_now():
    """手动重载全部宏缓存。"""
    if not normalize_hotkey_to_name(MACRO_RELOAD_KEY):
        print("[宏预编译] 未配置重载键，忽略重载请求")
        return
    _build_compiled_macro_cache_for_all_bindings("手动重载")


def _load_compiled_macro_lazy(trigger_key, filepath):
    """按需加载/编译单个宏。"""
    script = load_macro_from_file(filepath)
    compiled = compile_macro(script)
    return {
        "compiled_macro": compiled,
        "macro_name": os.path.basename(filepath),
        "macro_path": filepath,
        "loaded_at": time.strftime("%H:%M:%S"),
    }


def stop_and_wait_trigger_macros(trigger_key, wait_timeout_s=1.5):
    """停止触发键对应实例并等待退出。"""
    stop_macros_by_trigger(trigger_key, reason="再次按下-重启前停止")
    deadline = time.perf_counter() + max(0.0, wait_timeout_s)
    while has_live_macro_for_trigger(trigger_key):
        if time.perf_counter() >= deadline:
            return False
        time.sleep(0.01)
    return True


def stop_macros_by_trigger(trigger_key, reason="按键抬起停止"):
    """停止某个触发键启动的所有宏实例"""
    with active_macros_lock:
        macro_ids = list(trigger_to_macro_ids.get(trigger_key, set()))
        events = [active_macros[mid]["stop_event"] for mid in macro_ids if mid in active_macros]

    if macro_ids:
        print(f"[宏停止] {reason}：{trigger_key} -> 停止{len(macro_ids)}个宏实例")
    for ev in events:
        ev.set()


def stop_all_macros(reason="手动停止"):
    """停止全部正在运行的宏实例"""
    with active_macros_lock:
        events = [info["stop_event"] for info in active_macros.values()]
        count = len(events)

    if count:
        print(f"\n[宏停止] {reason}：停止全部宏实例（{count}个）")
    for ev in events:
        ev.set()


def toggle_global_trigger_pause():
    """切换全局触发暂停状态。"""
    global global_trigger_paused
    with active_macros_lock:
        global_trigger_paused = not global_trigger_paused
        paused_now = global_trigger_paused

    if paused_now:
        stop_all_macros("全局暂停键触发（等同ESC）")
        print("[全局暂停] 已启用：宏触发被禁止。再次按全局暂停键可恢复")
    else:
        print("[全局暂停] 已解除：宏触发恢复")


def is_global_trigger_paused():
    with active_macros_lock:
        return global_trigger_paused


def _pick_random_signed_range(down_value, up_value):
    """将 down/up 配置统一为 [lower, upper] 的随机区间，支持正负输入。"""
    lower, upper = -float(down_value), float(up_value)
    if lower > upper:
        lower, upper = upper, lower
    return random.uniform(lower, upper)


def calc_adjusted_delay_seconds(base_delay_seconds):
    """计算单步动作修正后的最终延迟（秒）。"""
    if not ENABLE_RANDOM_DELAY_ADJUST:
        return max(0.0, base_delay_seconds)

    base_delay_ms = max(0.0, base_delay_seconds * 1000.0)
    delta_ms = 0.0

    for rule in RANDOM_DELAY_ADJUST_RULES_BY_DELAY_MS:
        min_ms = rule.get("min", 0)
        max_ms = rule.get("max")
        if not (base_delay_ms >= min_ms and (max_ms is None or base_delay_ms < max_ms)):
            continue

        fixed_delta_ms = 0.0
        if "fixed_ms" in rule:
            fixed_delta_ms = float(rule["fixed_ms"])
        elif "fixed_ms_down" in rule or "fixed_ms_up" in rule:
            fixed_delta_ms = _pick_random_signed_range(
                rule.get("fixed_ms_down", 0), rule.get("fixed_ms_up", 0))

        percent_delta = 0.0
        if "percent" in rule:
            percent_delta = float(rule["percent"])
        elif "percent_down" in rule or "percent_up" in rule:
            percent_delta = _pick_random_signed_range(
                rule.get("percent_down", 0), rule.get("percent_up", 0))

        delta_ms = fixed_delta_ms + base_delay_ms * percent_delta / 100.0
        break

    return max(0.0, (base_delay_ms + delta_ms) / 1000.0)


def remove_comments(json_str):
    """移除JSON中的注释"""
    def replacer(match):
        s = match.group(0)
        return " " if s.startswith("/") else s
    return _COMMENT_PATTERN.sub(replacer, json_str)


def resolve_macro_file_path(path_hint, base_dir):
    """解析宏文件路径：支持当前目录、项目根目录与宏目录。"""
    if os.path.isabs(path_hint):
        return path_hint

    script_dir = os.path.dirname(os.path.abspath(__file__))
    search_roots = (
        base_dir, os.getcwd(), os.path.join(os.getcwd(), "宏"),
        script_dir, os.path.join(script_dir, "宏"),
    )

    for root in search_roots:
        candidate = os.path.abspath(os.path.join(root, path_hint))
        if os.path.exists(candidate):
            return candidate

    return os.path.abspath(os.path.join(base_dir, path_hint))


def load_macro_recursive(filepath, visited_files=None):
    """递归加载宏文件（支持import和loop）"""
    if visited_files is None:
        visited_files = set()

    abs_path = resolve_macro_file_path(filepath, os.getcwd())
    if abs_path in visited_files:
        print(f"[警告] 检测到宏文件循环引用，跳过：{filepath}")
        return []
    if not os.path.exists(abs_path):
        print(f"[错误] 宏文件不存在：{filepath}")
        return []

    visited_files.add(abs_path)
    base_dir = os.path.dirname(abs_path)
    final_script = []

    try:
        with open(abs_path, "r", encoding="utf-8") as f:
            raw_content = f.read()
        clean_content = remove_comments(raw_content)
        macro_data = json.loads(clean_content)

        for item in macro_data:
            if isinstance(item, list):
                if len(item) >= 2 and item[0] == "import":
                    sub_file = resolve_macro_file_path(item[1], base_dir)
                    print(f"[宏加载] 导入子宏：{sub_file}")
                    final_script.extend(load_macro_recursive(sub_file, visited_files))
                elif len(item) >= 3 and item[0] == "loop":
                    def process_loop_block(block):
                        processed = []
                        for sub_item in block:
                            if isinstance(sub_item, list):
                                if len(sub_item) >= 2 and sub_item[0] == "import":
                                    sub_file = resolve_macro_file_path(sub_item[1], base_dir)
                                    processed.extend(load_macro_recursive(sub_file, visited_files))
                                elif len(sub_item) >= 3 and sub_item[0] == "loop":
                                    sub_item[2] = process_loop_block(sub_item[2])
                                    processed.append(sub_item)
                                else:
                                    processed.append(sub_item)
                            else:
                                processed.append(sub_item)
                        return processed

                    item[2] = process_loop_block(item[2])
                    final_script.append(item)
                else:
                    final_script.append(item)
            else:
                final_script.append(item)

    except json.JSONDecodeError as e:
        print(f"[错误] 宏文件JSON格式错误（{filepath}）：{e}")
    except Exception as e:
        print(f"[错误] 加载宏文件失败（{filepath}）：{e}")
    finally:
        visited_files.discard(abs_path)

    return final_script


def load_macro_from_file(filepath):
    """加载宏文件入口"""
    print(f"[宏加载] 开始加载宏文件：{filepath}")
    macro_script = load_macro_recursive(filepath)
    print(f"[宏加载] 加载完成，总指令数：{len(macro_script)}")
    return macro_script


def compile_macro(script):
    """编译宏脚本为可执行指令（修复闭包陷阱）"""
    if not script:
        return []
    compiled = []
    compiled_actions = 0
    max_actions = MAX_COMPILED_ACTIONS

    def _reserve_actions(n):
        nonlocal compiled_actions
        compiled_actions += n
        if compiled_actions > max_actions:
            raise ValueError(f"编译动作数超过上限（{max_actions}），请降低宏内loop次数或拆分宏")

    def _parse(item):
        cmd = item[0]
        if cmd == "loop":
            block = []
            loop_body = item[2]
            loop_count = int(item[1])
            for _ in range(loop_count):
                for sub in loop_body:
                    block.extend(_parse(sub))
            return block

        delay_s, func = 0.0, None
        if cmd == "wait":
            delay_s = max(0.0, float(item[1]) / 1000.0)
            _reserve_actions(1)
            return [(None, delay_s, "wait", f"wait({float(item[1]):.1f})")]

        if cmd == "view":
            dx, dy, dur = item[1][0], item[1][1], max(item[2], 1)
            min_step_interval = 10
            max_steps = 100
            min_steps = 2
            target_steps = int(dur / min_step_interval)
            steps = max(min_steps, min(target_steps, max_steps))
            sx = int(dx / steps)
            sy = int(dy / steps)
            s_del = (dur / steps) / 1000.0
            _reserve_actions(steps)
            return [
                (lambda x=sx, y=sy: send_mouse_move(x, y), s_del, "view",
                 f"view(dx={sx},dy={sy},step={s_del * 1000:.1f})")
                for _ in range(steps)
            ]

        key_name = item[1]
        vk = get_vk(key_name)
        delay_s = max(0.0, float(item[2]) / 1000.0)
        action_desc = f"{cmd}({key_name})"

        if cmd == "kd":
            func = lambda v=vk: send_key_event(v, False)
        elif cmd == "ku":
            func = lambda v=vk: send_key_event(v, True)
        elif cmd == "md":
            is_left = (key_name.lower() == "left")
            func = lambda left=is_left: send_mouse_click(left, False)
        elif cmd == "mu":
            is_left = (key_name.lower() == "left")
            func = lambda left=is_left: send_mouse_click(left, True)

        if func:
            _reserve_actions(1)

        return [(func, delay_s, cmd, action_desc)] if func else []

    for step in script:
        compiled.extend(_parse(step))
    return compiled


# ==============================================================================
# 宏执行核心逻辑（支持执行中外部干预）
# ==============================================================================
def execute_macro_once(compiled_macro, stop_event, pause_event=None, macro_name="未知宏"):
    """
    执行编译后的宏
    核心改进：
    1. 每步执行后检查停止信号，支持中途终止
    2. 采用非独占式执行，允许物理鼠标/键盘干预
    3. 加入最大执行时间限制，防止死循环
    """
    if not compiled_macro:
        print(f"[宏执行] {macro_name} 无指令可执行")
        return

    if pause_event is None:
        pause_event = threading.Event()

    print(f"[宏执行] 开始执行 {macro_name}")
    trace_mode = str(TEST_TRACE_OUTPUT_MODE).lower()
    if trace_mode not in ("realtime", "final"):
        trace_mode = "final"
    if ENABLE_TEST_TRACE:
        mode_text = "逐行实时" if trace_mode == "realtime" else "执行后统一输出"
        print(f"[测试] 追踪已开启：{macro_name} 共{len(compiled_macro)}步 | 模式={mode_text}")
        if trace_mode == "realtime":
            print("\n          步数 |     指令     | 驱动发包(ms) | 实际等待(ms) | 补偿(ms) |  误差(ms) | 累计偏差(ms)")

    start_time = time.perf_counter()
    gc_was_enabled = gc.isenabled()

    reports = []
    target_total_time_s = 0.0
    total_driver_cost_ms = 0.0
    max_step_jitter_ms = -float("inf")
    min_step_jitter_ms = float("inf")
    paused_total_s = 0.0
    finished_elapsed_s = None
    finished_effective_elapsed_s = None

    # 累计误差控制器（PI 控制策略）
    class CumulativeLagController:
        """PI控制器 - 维护理想时间线并进行误差补偿"""

        def __init__(self, kp=0.15, ki=0.003, error_trigger_ms=2.0, error_target_ms=1.5,
                     integral_max=2.0, integral_decay=0.998, max_step_comp_pct=0.35,
                     min_step_delay_ms=0.3, process_noise=1e-6, measure_noise=2e-5):
            self.controller_name = "PI"
            self.accepts_driver_cost = False
            self.kp = kp
            self.ki = ki
            self.error_trigger = error_trigger_ms / 1000.0
            self.error_target = error_target_ms / 1000.0
            self.integral_max = integral_max
            self.integral_decay = integral_decay
            self.max_step_comp_pct = max_step_comp_pct
            self.min_step_delay = min_step_delay_ms / 1000.0

            # 一阶卡尔曼滤波器状态
            self._est = 0.0
            self._cov = 1.0
            self._q = process_noise
            self._r = measure_noise
            self._integral = 0.0
            self.total_compensated = 0.0
            self.total_recovered = 0.0
            self.compensation_count = 0

        def _kalman_update(self, measurement):
            """一阶卡尔曼滤波"""
            self._cov += self._q
            gain = self._cov / (self._cov + self._r)
            self._est += gain * (measurement - self._est)
            self._cov *= (1.0 - gain)
            return self._est

        def reset_state(self):
            self._est = 0.0
            self._cov = 1.0
            self._integral = 0.0

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

    # 前馈 + 比例延迟补偿控制器
    class FeedforwardDelayController:
        """前馈 + 比例控制器 - 直接测量并抵消驱动开销"""

        def __init__(self, window_size=FF_WINDOW_SIZE, warmup_steps=FF_WARMUP_STEPS,
                     scale=FF_SCALE, kp=FF_KP, dead_zone_ms=FF_DEAD_ZONE_MS,
                     max_comp_pct=FF_MAX_COMP_PCT, rate_limit_pct=FF_RATE_LIMIT_PCT,
                     anti_overshoot=FF_ANTI_OVERSHOOT, ema_alpha=FF_EMA_ALPHA,
                     min_step_delay_ms=LAG_MIN_STEP_DELAY_MS):
            self.controller_name = "Feedforward+P"
            self.accepts_driver_cost = True

            # 前馈状态
            self._window = deque(maxlen=window_size)
            self._avg_driver_cost = 0.0
            self._scale = scale
            self._warmup_steps = warmup_steps

            # 比例控制状态
            self._kp = kp
            self._dead_zone = dead_zone_ms / 1000.0
            self._ema_alpha = ema_alpha
            self._smoothed_error = 0.0

            # 防过冲状态
            self._rate_limit_pct = rate_limit_pct
            self._anti_overshoot = anti_overshoot
            self._prev_compensation = 0.0
            self._prev_error_sign = 0

            # 限幅
            self._max_comp_pct = max_comp_pct
            self._min_step_delay = min_step_delay_ms / 1000.0

            # 统计
            self.compensation_count = 0
            self.total_compensated = 0.0
            self.total_feedforward = 0.0
            self.total_proportional = 0.0

        def compute_compensation(self, driver_cost, current_time, ideal_time,
                                 next_delay, action_type=""):
            # Layer 0: 更新前馈估计
            self._window.append(driver_cost)
            if len(self._window) >= self._warmup_steps:
                self._avg_driver_cost = sum(self._window) / len(self._window)

            if action_type == "wait":
                return 0.0

            # Layer 1: 前馈补偿
            ff_comp = self._avg_driver_cost * self._scale if len(self._window) >= self._warmup_steps else 0.0

            # Layer 2: 比例修正（无积分）
            raw_error = current_time - ideal_time
            self._smoothed_error = (self._ema_alpha * raw_error +
                                    (1 - self._ema_alpha) * self._smoothed_error)

            if abs(self._smoothed_error) < self._dead_zone:
                p_comp = 0.0
            else:
                p_comp = self._kp * self._smoothed_error
                curr_sign = 1 if self._smoothed_error > 0 else (-1 if self._smoothed_error < 0 else 0)
                if (self._prev_error_sign != 0 and curr_sign != 0 and
                        curr_sign != self._prev_error_sign):
                    p_comp *= self._anti_overshoot
                self._prev_error_sign = curr_sign

            # 总补偿 = 前馈 + 比例
            total_comp = ff_comp + p_comp

            # Layer 3: 安全限幅
            max_comp = next_delay * self._max_comp_pct
            total_comp = max(-max_comp, min(total_comp, max_comp))

            max_delta = next_delay * self._rate_limit_pct
            delta = total_comp - self._prev_compensation
            if abs(delta) > max_delta:
                total_comp = self._prev_compensation + max_delta * (1 if delta > 0 else -1)
            self._prev_compensation = total_comp

            # 统计
            if abs(total_comp) > 1e-7:
                self.compensation_count += 1
                self.total_compensated += total_comp
                self.total_feedforward += ff_comp
                self.total_proportional += p_comp

            return total_comp

        def reset_state(self):
            self._smoothed_error = 0.0
            self._prev_compensation = 0.0
            self._prev_error_sign = 0

    # 初始化控制器
    lag_controllers = []
    if LAG_COMPENSATION_ENABLED:
        if LAG_USE_PI_CONTROLLER:
            lag_controllers.append(CumulativeLagController(
                kp=LAG_KP, ki=LAG_KI, error_trigger_ms=LAG_ERROR_TRIGGER_MS,
                error_target_ms=LAG_ERROR_TARGET_MS, integral_max=LAG_INTEGRAL_MAX,
                integral_decay=LAG_INTEGRAL_DECAY, max_step_comp_pct=LAG_MAX_STEP_COMP_PCT,
                min_step_delay_ms=LAG_MIN_STEP_DELAY_MS, process_noise=LAG_FILTER_PROCESS_NOISE,
                measure_noise=LAG_FILTER_MEASURE_NOISE,
            ))
        if LAG_USE_FEEDFORWARD:
            lag_controllers.append(FeedforwardDelayController())

    def wait_if_paused(reset_reason):
        """若处于暂停态，则等待恢复并重置误差时间线。"""
        nonlocal ideal_time, paused_total_s
        if not pause_event.is_set() or stop_event.is_set():
            return 0.0

        pause_start = time.perf_counter()
        while pause_event.is_set() and not stop_event.is_set():
            time.sleep(0.01)
        pause_duration = max(0.0, time.perf_counter() - pause_start)
        paused_total_s += pause_duration

        for controller in lag_controllers:
            if hasattr(controller, "reset_state"):
                controller.reset_state()

        ideal_time = time.perf_counter()

        if ENABLE_TEST_TRACE:
            pause_ms = pause_duration * 1000.0
            row = {"step": "PAUSE", "desc": f"pause({reset_reason})",
                   "driver_ms": 0.0, "wait_ms": pause_ms, "comp_ms": 0.0,
                   "error_ms": 0.0, "cum_bias_ms": 0.0}
            reports.append(row)
            if trace_mode == "realtime":
                print(f"[动作表] {'PAUSE':>5} | {row['desc']:<12} | {row['driver_ms']:>12.3f} | "
                      f"{row['wait_ms']:>12.3f} | {row['comp_ms']:>8.3f} | {row['error_ms']:>9.3f} | "
                      f"{row['cum_bias_ms']:>11.3f}")

        return pause_duration

    ideal_time = start_time

    if gc_was_enabled:
        gc.disable()

    try:
        for idx, (func, delay, action_type, action_desc) in enumerate(compiled_macro):
            wait_if_paused("start")

            if stop_event.is_set():
                print(f"[宏执行] {macro_name} 被手动终止（执行到第{idx + 1}步）")
                return False

            if time.perf_counter() - start_time > MAX_MACRO_EXEC_TIME:
                print(f"[宏执行] {macro_name} 执行超时（>{MAX_MACRO_EXEC_TIME}秒），强制终止")
                break

            # 执行动作
            action_exec_start = time.perf_counter()
            if func is not None:
                try:
                    func()
                except Exception as e:
                    print(f"[宏执行] 第{idx + 1}步执行出错：{e}")
                    continue
            action_exec_end = time.perf_counter()
            action_exec_cost = action_exec_end - action_exec_start
            current_time = action_exec_end

            if stop_event.is_set():
                print(f"[宏执行] {macro_name} 被手动终止（执行到第{idx + 1}步）")
                return False

            # 计算延迟
            adjusted_delay = max(0.0, delay) if action_type == "wait" else calc_adjusted_delay_seconds(delay)

            # 累计误差补偿
            compensation = 0.0
            for controller in lag_controllers:
                if getattr(controller, "accepts_driver_cost", False):
                    compensation += controller.compute_compensation(
                        driver_cost=action_exec_cost, current_time=current_time,
                        ideal_time=ideal_time, next_delay=adjusted_delay, action_type=action_type)
                else:
                    compensation += controller.compute_compensation(
                        current_time=current_time, ideal_time=ideal_time,
                        next_delay=adjusted_delay, action_type=action_type)

            ideal_time += adjusted_delay
            sleep_time = max(0.0, adjusted_delay - compensation)
            sleep_deadline = time.perf_counter() + sleep_time
            remaining_sleep = sleep_time
            paused_in_sleep_s = 0.0
            actual_sleep_s = 0.0

            while remaining_sleep > 0.0:
                if stop_event.is_set():
                    break
                if pause_event.is_set():
                    paused_in_sleep_s += wait_if_paused("sleep")
                    break
                now = time.perf_counter()
                remaining_sleep = max(0.0, sleep_deadline - now)
                if remaining_sleep <= 0.0:
                    break

                chunk = min(remaining_sleep, 0.001)
                before_sleep = time.perf_counter()
                precise_sleep(chunk, stop_event)
                slept = max(0.0, time.perf_counter() - before_sleep)
                actual_sleep_s += slept
                remaining_sleep = max(0.0, sleep_deadline - time.perf_counter())

            # 仅统计 sleep 引擎实际执行耗时，避免循环控制开销造成“系统性偏大”。
            actual_wait = max(0.0, actual_sleep_s)
            stats_wait_for_error = actual_wait
            if paused_in_sleep_s > 0.0:
                stats_wait_for_error = max(0.0, adjusted_delay - action_exec_cost)

            if stop_event.is_set():
                print(f"[宏执行] {macro_name} 被手动终止（执行到第{idx + 1}步）")
                return False

            if ENABLE_TEST_TRACE:
                step_error_ms = (action_exec_cost + stats_wait_for_error - adjusted_delay) * 1000.0
                cumulative_bias_ms = (time.perf_counter() - ideal_time) * 1000.0
                total_driver_cost_ms += action_exec_cost * 1000.0
                target_total_time_s += adjusted_delay
                max_step_jitter_ms = max(max_step_jitter_ms, step_error_ms)
                min_step_jitter_ms = min(min_step_jitter_ms, step_error_ms)

                row = {"step": idx + 1, "desc": action_desc,
                       "driver_ms": action_exec_cost * 1000.0, "wait_ms": stats_wait_for_error * 1000.0,
                       "comp_ms": compensation * 1000.0, "error_ms": step_error_ms,
                       "cum_bias_ms": cumulative_bias_ms}
                reports.append(row)

                if trace_mode == "realtime":
                    print(f"[动作表] {row['step']:>5} | {row['desc']:<12} | {row['driver_ms']:>12.3f} | "
                          f"{row['wait_ms']:>12.3f} | {row['comp_ms']:>8.3f} | {row['error_ms']:>9.3f} | "
                          f"{row['cum_bias_ms']:>11.3f}")

        # 执行结果判定
        if not stop_event.is_set():
            elapsed = time.perf_counter() - start_time
            extra_info = ""
            if lag_controllers:
                detail_parts = []
                for controller in lag_controllers:
                    if getattr(controller, "compensation_count", 0) <= 0:
                        continue
                    if controller.controller_name == "Feedforward+P":
                        detail_parts.append(
                            f"Feedforward+P: {controller.compensation_count}步, "
                            f"总补偿 {controller.total_compensated * 1000:.2f}ms, "
                            f"前馈 {controller.total_feedforward * 1000:.2f}ms, "
                            f"比例 {controller.total_proportional * 1000:.2f}ms")
                    else:
                        detail_parts.append(
                            f"PI: {controller.compensation_count}步, "
                            f"总补偿 {controller.total_compensated * 1000:.2f}ms")
                if detail_parts:
                    extra_info = " | 延迟修正: " + " ; ".join(detail_parts)

            effective_elapsed = max(0.0, elapsed - paused_total_s)
            finished_elapsed_s = elapsed
            finished_effective_elapsed_s = effective_elapsed
            print(f"[宏执行] {macro_name} 执行完成，总耗时：{elapsed:.4f}秒"
                  f"（去暂停={effective_elapsed:.4f}秒, 暂停={paused_total_s:.4f}秒）{extra_info}")
            return True
        return False

    finally:
        if gc_was_enabled:
            gc.enable()

        if ENABLE_TEST_TRACE and reports:
            action_reports = [r for r in reports if r["step"] != "PAUSE"]
            total_steps = len(action_reports)
            elapsed_s = finished_elapsed_s if finished_elapsed_s is not None else (time.perf_counter() - start_time)
            effective_elapsed_s = (finished_effective_elapsed_s if finished_effective_elapsed_s is not None
                                   else max(0.0, elapsed_s - paused_total_s))
            avg_driver_cost_ms = (total_driver_cost_ms / total_steps) if total_steps else 0.0
            final_abs_bias_ms = abs(action_reports[-1]["cum_bias_ms"]) if action_reports else 0.0

            if trace_mode == "final":
                print("\n          步数 |     指令     | 驱动发包(ms) | 实际等待(ms) | 补偿(ms) |  误差(ms) | 累计偏差(ms)")
                for r in reports:
                    step_label = f"{r['step']:>5}" if isinstance(r["step"], int) else f"{r['step']:>5}"
                    print(f"[动作表] {step_label} | {r['desc']:<12} | {r['driver_ms']:>12.3f} | "
                          f"{r['wait_ms']:>12.3f} | {r['comp_ms']:>8.3f} | {r['error_ms']:>9.3f} | "
                          f"{r['cum_bias_ms']:>11.3f}")

            print("-------------")
            print(f"宏目标总时间={target_total_time_s:.4f}s，宏执行总时长={effective_elapsed_s:.4f}s，"
                  f"指令总数={total_steps}，平均每步驱动开销={avg_driver_cost_ms:.3f}ms，"
                  f"最大单步时间抖动={max_step_jitter_ms:.3f}ms，"
                  f"最小单项修正抖动={min_step_jitter_ms:.3f}ms，"
                  f"最终绝对时间偏差={final_abs_bias_ms:.3f}ms")
            if paused_total_s > 0:
                print(f"[测试] 本次暂停总时长：{paused_total_s:.4f}s")

        # 安全释放所有修饰键和鼠标按键
        send_key_event(get_vk("shift"), True)
        send_key_event(get_vk("ctrl"), True)
        send_key_event(get_vk("alt"), True)
        send_mouse_click(True, True)
        send_mouse_click(False, True)


def trigger_macro(trigger_key, start_paused=False):
    """触发宏执行（支持多宏并发）"""
    cfg = resolve_macro_config(trigger_key)
    if not cfg:
        print(f"[宏触发] 未找到触发键配置：{trigger_key}")
        return

    cfg_mode = cfg.get("running_press_mode", "normal")
    allow_parallel = (cfg_mode == "parallel_trigger")

    ok, reason = try_reserve_trigger(trigger_key, allow_parallel=allow_parallel)
    if not ok:
        print(f"[宏触发] 跳过触发：{trigger_key}，{reason}")
        return

    filepath = cfg["macro_path"]
    repeat_count = cfg["repeat_count"]

    try:
        use_preload = bool(normalize_hotkey_to_name(MACRO_RELOAD_KEY))
        if use_preload:
            with macro_cache_lock:
                cached = compiled_macro_cache.get(trigger_key)
            if cached:
                compiled_macro = list(cached.get("compiled_macro", []))
                macro_name = cached.get("macro_name", os.path.basename(filepath))
                print(f"[宏触发] 使用预编译缓存：{trigger_key} -> {macro_name}")
            else:
                print(f"[宏触发] 预编译缓存缺失，回退即时编译：{trigger_key}")
                cached = _load_compiled_macro_lazy(trigger_key, filepath)
                compiled_macro = cached["compiled_macro"]
                macro_name = cached["macro_name"]
                with macro_cache_lock:
                    compiled_macro_cache[trigger_key] = cached
        else:
            cached = _load_compiled_macro_lazy(trigger_key, filepath)
            compiled_macro = cached["compiled_macro"]
            macro_name = cached["macro_name"]
    except Exception as e:
        release_starting_trigger(trigger_key)
        print(f"[宏触发] 加载/编译宏失败：{e}")
        return

    macro_id = next_macro_id()
    stop_event = threading.Event()
    pause_event = threading.Event()
    if start_paused:
        pause_event.set()

    def macro_worker():
        boost_thread_priority()
        if ENABLE_PRO_AUDIO and not apply_pro_audio(True):
            print(">>> [系统] ProAudio 注入失败，继续使用普通线程调度")
        total_cores = os.cpu_count() or 4
        target_core = total_cores - 2
        set_thread_affinity(target_core)
        try:
            if repeat_count > 1:
                print(f"[宏执行] {macro_name} 将重复执行 {repeat_count} 次")
                for i in range(repeat_count):
                    if stop_event.is_set():
                        break
                    print(f"[宏执行] {macro_name} 第 {i + 1}/{repeat_count} 次")
                    execute_macro_once(compiled_macro, stop_event, pause_event, macro_name)
            else:
                execute_macro_once(compiled_macro, stop_event, pause_event, macro_name)
        finally:
            unregister_active_macro(macro_id)

    worker = threading.Thread(target=macro_worker, daemon=True)
    register_active_macro(macro_id, trigger_key, macro_name, worker, stop_event, pause_event)
    release_starting_trigger(trigger_key)
    worker.start()

    print(f"[宏触发] 已启动实例#{macro_id}: key={trigger_key}, repeat={repeat_count}, "
          f"press={cfg['running_press_mode']}")
    if start_paused:
        print(f"[宏暂停] 首次按下仅预加载：{trigger_key}（抬起后开始执行）")


def handle_trigger_release(trigger_key):
    """按键/侧键抬起时，按配置决定是否停止对应宏实例"""
    cfg = resolve_macro_config(trigger_key)
    if not cfg:
        return
    mode = cfg["running_press_mode"]

    if mode == "hold_pause":
        resume_macros_by_trigger(trigger_key, reason="按住暂停模式-抬起继续")
    elif mode == "release_pause":
        pause_macros_by_trigger(trigger_key, reason="release_pause-抬起暂停")
    elif mode == "release_stop":
        stop_macros_by_trigger(trigger_key, reason="release_stop-抬起停止")


def handle_trigger_press(trigger_key):
    """按键/侧键按下时，综合配置处理触发/暂停/继续。"""
    cfg = resolve_macro_config(trigger_key)
    if not cfg:
        return

    if is_global_trigger_paused():
        print(f"[宏触发] 已被全局暂停拦截：{trigger_key}")
        return

    mode = cfg["running_press_mode"]

    if mode == "release_pause":
        if has_paused_macro_for_trigger(trigger_key):
            resume_macros_by_trigger(trigger_key, reason="按下触发键继续")
        elif has_live_macro_for_trigger(trigger_key):
            print(f"[宏触发] 宏已在运行：{trigger_key}")
        else:
            trigger_macro(trigger_key)
        return

    if mode == "pause_resume":
        if has_live_macro_for_trigger(trigger_key):
            if has_paused_macro_for_trigger(trigger_key):
                resume_macros_by_trigger(trigger_key, reason="再次按下继续")
            else:
                pause_macros_by_trigger(trigger_key, reason="再次按下暂停")
        return

    if mode == "hold_pause":
        if has_live_macro_for_trigger(trigger_key):
            pause_macros_by_trigger(trigger_key, reason="按住暂停模式-按下暂停")
        else:
            trigger_macro(trigger_key, start_paused=True)
        return

    if mode == "stop_restart":
        if has_live_macro_for_trigger(trigger_key):
            print(f"[宏触发] stop_restart：先停止并等待旧实例退出 -> {trigger_key}")
            if not stop_and_wait_trigger_macros(trigger_key):
                print(f"[宏触发] stop_restart 超时：旧实例未及时退出，取消本次重启 -> {trigger_key}")
                return
        trigger_macro(trigger_key)
        return

    if mode == "parallel_trigger":
        live_count = len(_collect_live_macro_infos(trigger_key))
        if live_count > 0:
            print(f"[宏触发] parallel_trigger：新增并发实例（当前已有{live_count}个）")
        trigger_macro(trigger_key)
        return

    if mode == "release_stop" and has_live_macro_for_trigger(trigger_key):
        print(f"[宏触发] release_stop：宏已在运行，等待抬起停止 -> {trigger_key}")
        return

    trigger_macro(trigger_key)


# ==============================================================================
# 鼠标按键监听回调
# ==============================================================================
def on_mouse_click(x, y, button, pressed):
    """鼠标按键点击回调（处理侧键按下与抬起）"""
    key_str = "mouse_x1" if button == mouse.Button.x1 else ("mouse_x2" if button == mouse.Button.x2 else None)
    if key_str is None:
        return

    if not pressed:
        handle_trigger_release(key_str)
        return

    if key_str in MACRO_BINDINGS:
        macro_path = MACRO_BINDINGS[key_str]
        print(f"\n[鼠标监听] 检测到侧键{key_str}，触发宏：{macro_path}")
        handle_trigger_press(key_str)


# ==============================================================================
# 键盘监听逻辑
# ==============================================================================
def on_key_press(key):
    """按键按下回调"""
    try:
        key_str = get_pressed_key_name(key)
        if not key_str:
            return

        if is_hotkey_match(GLOBAL_PAUSE_TOGGLE_KEY, key_str):
            toggle_global_trigger_pause()
            return

        if is_hotkey_match(STOP_MACRO_KEY, key_str):
            stop_all_macros(f"{key_str}触发")
            return

        if is_reload_hotkey(key_str):
            print(f"\n[按键监听] 检测到重载键 {key_str}，开始重新加载宏配置")
            reload_macro_cache_now()
            return

        if key_str in pressed_trigger_keys:
            return
        pressed_trigger_keys.add(key_str)

        if key_str in MACRO_BINDINGS:
            macro_path = MACRO_BINDINGS[key_str]
            print(f"\n[按键监听] 检测到{key_str}键，触发宏：{macro_path}")
            handle_trigger_press(key_str)
    except Exception as e:
        print(f"[按键监听] 出错：{e}")


def on_key_release(key):
    """按键抬起回调"""
    try:
        key_str = get_pressed_key_name(key)
        if not key_str:
            return
        pressed_trigger_keys.discard(key_str)
        handle_trigger_release(key_str)
    except Exception as e:
        print(f"[按键监听] 抬起处理出错：{e}")


# ==============================================================================
# 启动键盘+鼠标双监听
# ==============================================================================
def start_listeners():
    """启动键盘+鼠标双监听线程"""
    boost_process_priority()
    if ENABLE_REALTIME_PRIORITY:
        ok = apply_realtime_priority(True)
        print(f"[系统] 进程实时优先级: {'已开启' if ok else '开启失败，保持当前'}")

    auto_warmup()
    init_interception_backend()
    preload_macros_if_enabled()

    print("=" * 50)
    print("[系统] 通用宏执行框架已启动（鼠标侧键+键盘双触发）")
    print("[系统] 宏绑定列表：")
    for key, path in MACRO_BINDINGS.items():
        cfg = resolve_macro_config(key)
        if cfg:
            print(f"  {key} -> {path} | repeat={cfg['repeat_count']} | press={cfg['running_press_mode']}")
        else:
            print(f"  {key} -> {path}")

    reload_key_name = normalize_hotkey_to_name(MACRO_RELOAD_KEY)
    if reload_key_name:
        print(f"[系统] 宏重载键：{reload_key_name}（按下后将重新加载并编译全部绑定宏）")
    else:
        print("[系统] 宏重载键：未配置（按触发键时即时加载/编译）")

    stop_key_name = normalize_hotkey_to_name(STOP_MACRO_KEY) or str(STOP_MACRO_KEY)
    global_pause_key_name = normalize_hotkey_to_name(GLOBAL_PAUSE_TOGGLE_KEY) or str(GLOBAL_PAUSE_TOGGLE_KEY)
    print(f"[系统] 按{stop_key_name}终止所有正在执行的宏 | 按{global_pause_key_name}切换全局暂停/恢复")

    if USE_INTERCEPTION:
        print(f"[系统] 输入后端: keyboard={'Interception' if USE_INTERCEPTION_KEYBOARD else 'SendInput'}, "
              f"mouse={'Interception' if USE_INTERCEPTION_MOUSE else 'SendInput'}")
    else:
        print("[系统] 输入后端: keyboard=SendInput, mouse=SendInput")

    print(f"[系统] 调度优化: realtime={'on' if ENABLE_REALTIME_PRIORITY else 'off'}, "
          f"pro_audio={'on' if ENABLE_PRO_AUDIO else 'off'}")

    if ENABLE_RANDOM_DELAY_ADJUST:
        print("[系统] 动作延迟随机修正：已启用")
        print(f"[系统] 动作延迟分档规则：{RANDOM_DELAY_ADJUST_RULES_BY_DELAY_MS}")

    if LAG_COMPENSATION_ENABLED:
        comp_info = ""
        if LAG_USE_PI_CONTROLLER:
            comp_info += " | PI控制器：已启用"
        if LAG_USE_FEEDFORWARD:
            comp_info += " | 前馈+比例控制器：已启用"
        print("[系统] 误差补偿" + (comp_info or " | 未启用任何控制器！"))
    else:
        print("[系统] 累计误差补偿：关闭")
    print("=" * 50)

    key_listener = keyboard.Listener(on_press=on_key_press, on_release=on_key_release)
    key_listener.daemon = True
    key_listener.start()

    mouse_listener = mouse.Listener(on_click=on_mouse_click)
    mouse_listener.daemon = True
    mouse_listener.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[系统] 框架被手动关闭")
    finally:
        stop_all_macros("系统退出")
        key_listener.stop()
        mouse_listener.stop()


# ==============================================================================
# 主函数
# ==============================================================================
if __name__ == "__main__":
    if not os.path.exists("宏"):
        os.makedirs("宏")
        print("[系统] 已创建「宏」文件夹，请将宏文件放入该目录")

    start_listeners()
