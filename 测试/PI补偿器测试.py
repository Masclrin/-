# -*- coding: utf-8 -*-
"""
智能动态PI补偿器 - 宏执行框架高精度误差控制

核心特性：
1. 自适应增益：根据误差幅度动态调整Kp/Ki
2. 导数预测：基于误差变化率预测未来趋势
3. 多级响应：不同误差区间采用不同策略
4. 智能积分：累计误差越低，积分作用越弱
5. 双向补偿：支持正负补偿（延迟/提早）
6. 目标稳定：累计偏差稳定在±0.3ms内
"""

import time
import math
from typing import Optional, Tuple

class SmartPIController:
    """
    智能动态PI控制器
    
    与传统PI控制器的区别：
    ┌─────────────────────────────────────────────────────────────────────┐
    │ 传统PI:  U(t) = Kp*e(t) + Ki*∫e(t)dt                               │
    │                                                                     │
    │ 智能PI:  U(t) = Kp(e,t)*e(t) + Ki(e,t)*∫e(t)dt + Kd(e,t)*de/dt    │
    │          其中 Kp, Ki, Kd 均随误差状态动态调整                        │
    │                                                                     │
    │ 关键改进：                                                          │
    │ 1. 误差越大 → Kp越大（快速拉回）                                    │
    │ 2. 累计误差越小 → Ki越小（避免过冲）                                │
    │ 3. 误差在减小 → 降低增益（防止震荡）                                │
    │ 4. 误差在增大 → 增大增益（加强响应）                                │
    │ 5. 误差接近目标 → 智能衰减积分项                                     │
    └─────────────────────────────────────────────────────────────────────┘
    """
    
    def __init__(
        self,
        # 基础参数（用于中等误差时的增益）
        base_kp: float = 0.3,
        base_ki: float = 0.08,
        base_kd: float = 0.15,
        
        # 目标与阈值
        target_error_ms: float = 0.0,        # 目标误差（通常为0）
        stability_band_ms: float = 0.3,     # 稳定区间±0.3ms
        compensation_threshold_ms: float = 0.1,  # 开始补偿的阈值
        
        # 动态增益范围
        kp_range: Tuple[float, float] = (0.1, 0.8),   # Kp动态范围
        ki_range: Tuple[float, float] = (0.01, 0.25), # Ki动态范围  
        kd_range: Tuple[float, float] = (0.05, 0.4),  # Kd动态范围
        
        # 积分限制
        integral_max: float = 3.0,           # 积分上限
        integral_decay_threshold: float = 0.5,  # 低于此误差时开始衰减积分
        
        # 补偿限制
        max_compensation_pct: float = 0.15,  # 最大单步补偿百分比（15%）
        
        # 卡尔曼滤波
        process_noise: float = 1e-5,
        measurement_noise: float = 3e-4,
        
        # 预热参数
        warmup_steps: int = 5
    ):
        # 基础增益
        self.base_kp = base_kp
        self.base_ki = base_ki
        self.base_kd = base_kd
        
        # 阈值参数
        self.target_error = target_error_ms / 1000.0  # 转换为秒
        self.stability_band = stability_band_ms / 1000.0
        self.compensation_threshold = compensation_threshold_ms / 1000.0
        
        # 动态增益范围
        self.kp_min, self.kp_max = kp_range
        self.ki_min, self.ki_max = ki_range
        self.kd_min, self.kd_max = kd_range
        
        # 积分限制
        self.integral_max = integral_max
        self.integral_decay_threshold = integral_decay_threshold
        
        # 补偿限制
        self.max_compensation_pct = max_compensation_pct
        
        # 卡尔曼滤波器状态
        self._kalman_est = 0.0
        self._kalman_cov = 1.0
        self._kalman_q = process_noise
        self._kalman_r = measurement_noise
        
        # PI控制器状态
        self._integral = 0.0
        self._prev_error = 0.0
        self._error_history = []  # 误差历史用于趋势分析
        self._history_window = 8   # 趋势分析窗口
        
        # 统计信息
        self._step_count = 0
        self._compensation_count = 0
        self._total_positive_comp = 0.0  # 缩短延迟的总补偿量
        self._total_negative_comp = 0.0  # 延长延迟的总补偿量（负补偿）
        
        # 预热
        self.warmup_steps = warmup_steps
        self._in_warmup = True
        
        # 调试信息
        self._debug_enabled = False
        
    @property
    def is_warmed_up(self) -> bool:
        """是否已完成预热"""
        return not self._in_warmup
    
    @property
    def current_integral(self) -> float:
        """当前积分值（秒）"""
        return self._integral
    
    @property
    def stats(self) -> dict:
        """获取统计信息"""
        return {
            'step_count': self._step_count,
            'compensation_count': self._compensation_count,
            'total_positive_comp_ms': self._total_positive_comp * 1000,
            'total_negative_comp_ms': self._total_negative_comp * 1000,
            'current_integral_ms': self._integral * 1000,
            'in_warmup': self._in_warmup
        }
    
    def _kalman_update(self, measurement: float) -> float:
        """
        一阶卡尔曼滤波：平滑误差观测，滤除瞬时抖动
        
        Args:
            measurement: 原始误差观测值（秒）
            
        Returns:
            平滑后的误差估计值
        """
        # 预测
        self._kalman_cov += self._kalman_q
        
        # 更新
        gain = self._kalman_cov / (self._kalman_cov + self._kalman_r)
        self._kalman_est += gain * (measurement - self._kalman_est)
        self._kalman_cov *= (1.0 - gain)
        
        return self._kalman_est
    
    def _compute_error_magnitude(self, error: float) -> float:
        """
        计算误差幅度等级（用于动态增益调整）
        
        Returns:
            误差等级 0.0~1.0+（越大表示误差越严重）
        """
        abs_error = abs(error)
        
        # 根据稳定区间定义误差等级
        if abs_error <= self.stability_band:
            return 0.0  # 在稳定区间内
        elif abs_error <= self.compensation_threshold:
            return abs_error / self.compensation_threshold  # 0~1
        elif abs_error <= self.stability_band * 5:
            return 1.0 + (abs_error - self.compensation_threshold) / (self.stability_band * 4)
        else:
            # 误差很大时，限制增长
            return 2.0 + math.log1p(abs_error - self.stability_band * 5) * 0.5
    
    def _compute_error_trend(self) -> float:
        """
        计算误差变化趋势
        
        Returns:
            趋势系数：
            > 0: 误差在增大（需要加强补偿）
            = 0: 误差稳定
            < 0: 误差在减小（可以降低补偿）
        """
        if len(self._error_history) < 3:
            return 0.0
        
        # 使用最近几次误差的变化率
        recent_errors = self._error_history[-5:]
        
        # 计算平均变化率
        changes = [recent_errors[i] - recent_errors[i-1] for i in range(1, len(recent_errors))]
        avg_change = sum(changes) / len(changes) if changes else 0.0
        
        # 归一化到 -1 ~ 1 范围
        threshold = self.compensation_threshold * 0.1
        trend = max(-1.0, min(1.0, avg_change / threshold))
        
        return trend
    
    def _compute_dynamic_gains(self, error: float) -> Tuple[float, float, float]:
        """
        根据当前误差状态动态计算增益
        
        Args:
            error: 当前平滑后的误差（秒）
            
        Returns:
            (Kp, Ki, Kd) 动态调整后的增益
        """
        # 计算误差幅度等级
        magnitude = self._compute_error_magnitude(error)
        
        # 基础增益
        kp = self.base_kp
        ki = self.base_ki
        kd = self.base_kd
        
        # ── 1. 根据误差幅度调整增益 ──
        if magnitude > 0:
            # 误差越大，比例增益越高（快速拉回）
            kp_scale = 1.0 + magnitude * 1.5  # 大误差时Kp可增大至2.5倍
            kp = max(self.kp_min, min(self.kp_max, kp * kp_scale))
            
            # 积分增益随误差增大而增加
            ki_scale = 1.0 + magnitude * 1.2
            ki = max(self.ki_min, min(self.ki_max, ki * ki_scale))
            
            # 导数增益也随误差增大
            kd_scale = 1.0 + magnitude * 0.8
            kd = max(self.kd_min, min(self.kd_max, kd * kd_scale))
        
        # ── 2. 根据误差趋势调整增益 ──
        trend = self._compute_error_trend()
        
        if trend > 0.2:
            # 误差在增大（发散），加强响应
            kp *= 1.3
            ki *= 1.4
            kd *= 1.2
        elif trend < -0.2:
            # 误差在减小（收敛），降低响应防止过冲
            kp *= 0.7
            ki *= 0.6
            kd *= 0.8
        
        # ── 3. 根据累计误差绝对值调整积分 ──
        # 核心改进：累计误差越小，积分作用越弱
        abs_integral = abs(self._integral)
        if abs_integral < self.integral_decay_threshold:
            # 在小积分区间，智能衰减
            decay_factor = abs_integral / self.integral_decay_threshold
            ki *= (0.3 + 0.7 * decay_factor)  # Ki降低至30%~100%
        
        return kp, ki, kd
    
    def _update_error_history(self, error: float):
        """更新误差历史"""
        self._error_history.append(error)
        if len(self._error_history) > self._history_window:
            self._error_history.pop(0)
    
    def compute_compensation(
        self, 
        current_time: float, 
        ideal_time: float, 
        next_delay: float,
        action_type: str = ""
    ) -> float:
        """
        计算补偿量（核心方法）
        
        Args:
            current_time: 当前实际时间
            ideal_time: 理想时间线（不含补偿）
            next_delay: 下一动作的延迟（秒）
            action_type: 动作类型（"wait"等）
            
        Returns:
            补偿量（秒）：
            > 0: 缩短睡眠时间（加速）
            < 0: 延长睡眠时间（减速，负补偿）
            = 0: 无补偿
        """
        self._step_count += 1
        
        # ── 预热阶段 ──
        if self._in_warmup:
            if self._step_count >= self.warmup_steps:
                self._in_warmup = False
                if self._debug_enabled:
                    print(f"[SmartPI] 预热完成，已执行{self._step_count}步")
            return 0.0
        
        # ── 计算原始误差 ──
        raw_error = current_time - ideal_time
        smoothed_error = self._kalman_update(raw_error)
        
        # ── wait指令不参与延迟补偿 ──
        if action_type == "wait":
            # 但仍更新历史和误差追踪
            self._update_error_history(smoothed_error)
            self._prev_error = smoothed_error
            return 0.0
        
        # ── 在稳定区间内 ──
        if abs(smoothed_error) <= self.stability_band:
            # 轻微衰减积分，不做补偿
            self._integral *= 0.98
            self._update_error_history(smoothed_error)
            self._prev_error = smoothed_error
            return 0.0
        
        # ── 误差未达到补偿阈值 ──
        if abs(smoothed_error) <= self.compensation_threshold:
            # 小误差，智能衰减积分
            self._integral *= 0.95
            self._update_error_history(smoothed_error)
            self._prev_error = smoothed_error
            return 0.0
        
        # ── 计算动态增益 ──
        kp, ki, kd = self._compute_dynamic_gains(smoothed_error)
        
        # ── PI(D) 控制计算 ──
        # 误差（相对于目标）
        error_from_target = smoothed_error - self.target_error
        
        # 比例项
        p_term = kp * error_from_target
        
        # 积分项
        self._integral += error_from_target
        # 双向积分限幅
        self._integral = max(-self.integral_max, min(self.integral_max, self._integral))
        i_term = ki * self._integral
        
        # 导数项（预测误差变化）
        derivative = smoothed_error - self._prev_error if self._step_count > 1 else 0.0
        d_term = kd * derivative
        
        # 总补偿
        total_compensation = p_term + i_term + d_term
        
        # ── 安全限幅 ──
        max_comp = next_delay * self.max_compensation_pct
        compensation = max(-max_comp, min(max_comp, total_compensation))
        
        # ── 更新状态 ──
        self._update_error_history(smoothed_error)
        self._prev_error = smoothed_error
        
        # 统计
        if compensation != 0.0:
            self._compensation_count += 1
            if compensation > 0:
                self._total_positive_comp += compensation
            else:
                self._total_negative_comp += abs(compensation)
        
        # 调试输出
        if self._debug_enabled and self._step_count % 50 == 0:
            print(
                f"[SmartPI] Step {self._step_count}: "
                f"error={smoothed_error*1000:.3f}ms, "
                f"Kp={kp:.3f}, Ki={ki:.3f}, Kd={kd:.3f}, "
                f"comp={compensation*1000:.3f}ms, "
                f"integral={self._integral*1000:.3f}ms"
            )
        
        return compensation  # 正=缩短睡眠，负=延长睡眠
    
    def reset(self):
        """重置控制器状态"""
        self._kalman_est = 0.0
        self._kalman_cov = 1.0
        self._integral = 0.0
        self._prev_error = 0.0
        self._error_history.clear()
        self._step_count = 0
        self._compensation_count = 0
        self._total_positive_comp = 0.0
        self._total_negative_comp = 0.0
        self._in_warmup = True
        
    def enable_debug(self, enable: bool = True):
        """启用/禁用调试输出"""
        self._debug_enabled = enable


class UltraPreciseController(SmartPIController):
    """
    超高精度控制器 - 针对±0.3ms目标的优化版本
    
    额外特性：
    1. 更精细的增益调度
    2. 更保守的补偿以避免过冲
    3. 额外的平滑处理
    """
    
    def __init__(self, **kwargs):
        # 优化后的默认参数
        defaults = {
            'base_kp': 0.25,
            'base_ki': 0.06,
            'base_kd': 0.12,
            'target_error_ms': 0.0,
            'stability_band_ms': 0.3,      # 目标：±0.3ms
            'compensation_threshold_ms': 0.15,  # 0.15ms开始补偿
            'kp_range': (0.08, 0.6),
            'ki_range': (0.005, 0.18),
            'kd_range': (0.03, 0.35),
            'integral_max': 2.5,
            'integral_decay_threshold': 0.4,
            'max_compensation_pct': 0.12,  # 单步最多12%
            'process_noise': 5e-6,
            'measurement_noise': 1e-4,
            'warmup_steps': 8
        }
        
        # 合并用户参数
        for key, value in defaults.items():
            if key not in kwargs:
                kwargs[key] = value
        
        super().__init__(**kwargs)
        
        # 额外平滑参数
        self._smoothing_alpha = 0.3  # 误差平滑系数
        self._smoothed_compensation = 0.0
        
    def compute_compensation(self, current_time: float, ideal_time: float, 
                           next_delay: float, action_type: str = "") -> float:
        """
        重写的补偿计算，增加额外的平滑处理
        """
        # 调用父类计算基础补偿
        raw_comp = super().compute_compensation(
            current_time, ideal_time, next_delay, action_type
        )
        
        if action_type == "wait":
            return raw_comp
        
        # 对补偿进行平滑，避免剧烈变化
        self._smoothed_compensation = (
            self._smoothing_alpha * raw_comp + 
            (1 - self._smoothing_alpha) * self._smoothed_compensation
        )
        
        # 再次限幅（双重保护）
        max_comp = next_delay * self.max_compensation_pct * 0.8
        final_comp = max(-max_comp, min(max_comp, self._smoothed_compensation))
        
        return final_comp


# =============================================================================
# 便捷工厂函数
# =============================================================================

def create_controller(
    mode: str = "smart",
    target_ms: float = 0.3,
    **kwargs
) -> SmartPIController:
    """
    创建补偿控制器
    
    Args:
        mode: 控制器模式
              - "smart": 智能动态PI控制器
              - "ultra": 超高精度控制器（针对±0.3ms优化）
              - "fast": 快速响应控制器（误差>1ms时更激进）
        target_ms: 目标稳定精度（ms）
        **kwargs: 其他参数传递给控制器
        
    Returns:
        SmartPIController实例
    """
    if mode == "ultra":
        controller = UltraPreciseController(
            stability_band_ms=target_ms,
            **kwargs
        )
    elif mode == "fast":
        kwargs.setdefault('kp_range', (0.15, 1.0))
        kwargs.setdefault('ki_range', (0.02, 0.35))
        kwargs.setdefault('max_compensation_pct', 0.2)
        controller = SmartPIController(
            stability_band_ms=target_ms,
            **kwargs
        )
    else:  # smart
        controller = SmartPIController(
            stability_band_ms=target_ms,
            **kwargs
        )
    
    return controller


# =============================================================================
# 测试代码
# =============================================================================
if __name__ == "__main__":
    print("=" * 70)
    print("智能PI补偿器测试")
    print("=" * 70)
    
    # 创建控制器
    controller = create_controller("ultra", target_ms=0.3)
    controller.enable_debug(True)
    
    # 模拟执行
    start_time = time.perf_counter()
    ideal_time = start_time
    
    errors = []  # 记录误差用于分析
    
    print(f"\n目标: 累计偏差稳定在 ±0.3ms 内")
    print(f"模拟100步执行...\n")
    
    for step in range(100):
        current_time = time.perf_counter()
        
        # 模拟延迟（有意引入一些延迟波动）
        if step < 10:
            # 预热阶段
            delay = 0.010  # 10ms
        elif step < 30:
            # 模拟系统开销导致的延迟累积
            delay = 0.010
            current_time += 0.003  # +3ms开销
        elif step < 50:
            # 继续累积
            delay = 0.010
            current_time += 0.002
        elif step < 70:
            # 误差开始收敛
            delay = 0.010
            current_time += 0.001
        else:
            # 正常执行，偶有小抖动
            delay = 0.010
            import random
            current_time += random.uniform(-0.0002, 0.0005)
        
        # 计算补偿
        compensation = controller.compute_compensation(
            current_time, ideal_time, delay, ""
        )
        
        # 更新理想时间线
        ideal_time += delay
        
        # 记录误差
        error = (current_time - ideal_time) * 1000  # ms
        errors.append(error)
        
        # 每20步输出一次统计
        if (step + 1) % 20 == 0:
            recent_errors = errors[-20:]
            abs_errors = [abs(e) for e in recent_errors]
            print(
                f"Steps {step-19}~{step+1}: "
                f"平均误差={sum(recent_errors)/20:.3f}ms, "
                f"最大={max(recent_errors):.3f}ms, "
                f"最小={min(recent_errors):.3f}ms, "
                f"稳定率={sum(1 for e in abs_errors if e <= 0.3)/20*100:.0f}%"
            )
    
    print(f"\n最终统计:")
    print(f"  总步数: {controller.stats['step_count']}")
    print(f"  补偿次数: {controller.stats['compensation_count']}")
    print(f"  正补偿总量: {controller.stats['total_positive_comp_ms']:.3f}ms")
    print(f"  负补偿总量: {controller.stats['total_negative_comp_ms']:.3f}ms")
    
    # 检查最终误差
    final_errors = errors[-20:]
    final_abs_errors = [abs(e) for e in final_errors]
    within_target = sum(1 for e in final_abs_errors if e <= 0.3)
    print(f"  最后20步在±0.3ms内的比例: {within_target}/20 = {within_target/20*100:.0f}%")
    
    print("=" * 70)