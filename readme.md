# 多窗口后台模拟点击器 0.3 By Ducky錡 + DeepSeek R1 @52pojie

一个功能强大的Windows后台点击自动化工具，支持多窗口、多模式的自动点击功能。

![Logo](./软件截图.jpg)

## 主要功能

### 1. 多窗口支持
- 支持同时对多个窗口进行自动点击
- 可以分别设置每个窗口的点击参数
- 支持窗口的精确定位和控件识别

### 2. 点击模式
提供4种不同的点击模式，适应不同场景：
- 模式1：使用控件句柄直接发送消息（最稳定）
- 模式2：通过窗口查找再点击
- 模式3：使用全局模拟点击
- 模式4：使用SendInput方式点击

### 3. 点击类型
支持多种点击类型：
- 左键单击
- 右键单击
- 左键双击
- 右键双击
- 左键双击(2秒) - 两次点击间隔2秒
- 右键双击(2秒) - 两次点击间隔2秒

### 4. 精确定位
- 使用控件相对坐标，确保点击位置准确
- 自动识别窗口和控件信息
- 支持实时显示鼠标位置信息

### 5. 定时控制
- 可设置点击间隔时间
- 支持随机延迟，避免固定间隔
- 每个任务可独立设置时间参数

### 6. 便捷操作
- F3：快速添加点击任务
- F9：开始/停止所有任务
- Delete：删除选中的任务
- 双击：编辑任务详细信息

### 7. 其他功能
- 详细的日志记录，方便排查问题
- 窗口置顶功能
- 任务导入导出
- 界面优化，操作直观

## 使用方法

1. 添加任务：
   - 将鼠标移动到目标位置
   - 按F3添加新任务
   - 或通过界面手动添加

2. 设置任务：
   - 双击任务列表中的项目进行编辑
   - 设置点击模式、类型、间隔等参数
   - 可随时启用/禁用单个任务

3. 开始运行：
   - 按F9开始执行所有已启用的任务
   - 再次按F9停止运行
   - 运行时可实时查看任务状态

## 注意事项

1. 部分窗口可能需要管理员权限
2. 不同点击模式适用于不同场景，建议逐个尝试
3. 使用时注意遵守相关法律法规
4. 建议先在测试环境验证配置

## 更新日志

### v1.0
- 初始版本发布
- 支持基本的多窗口点击功能

### v1.1
- 添加多种点击模式
- 优化界面布局
- 添加详细日志功能

### v1.2
- 新增双击功能
- 添加2秒间隔双击选项
- 优化坐标计算方式
- 改进任务编辑界面

## 关于作者

作者：Ducky錡
