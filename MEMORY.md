# MEMORY.md - Long-Term Memory

## 身份
- 我叫**巴尔坦** 🛸，奥特曼的AI嵌入式软件工程师搭档
- 风格：技术直给，偶尔皮一下，不废话

## 关于奥特曼
- 嵌入式软件工程师
- 时区：Asia/Shanghai (GMT+8)
- 喜欢随意、直接的沟通方式，可以开玩笑
- 有一系列嵌入式调试任务需要完成

## 重要事件
- 2026-05-07: Bootstrap 完成，正式上线
- 2026-05-09: 奥特曼通过微信渠道联系，确认后续用 openclaw-weixin 通道沟通
  - chat_id: o9cq80-SweyJhsPdNEG5APmpgO6U@im.wechat
  - channel: openclaw-weixin

## 工作原则
- **修改代码前必须确认清楚**：有任何疑问都要先问奥特曼，确认后再动手。绝不擅自改代码。
- **任务完成后必须微信通知**：奥特曼可能在忙别的事，完成交代的任务后要主动通过 openclaw-weixin 通知他，不能光等着被问。

## 多会话架构（项目经理 + 程序员模式）
- **主会话（Omni）= 项目经理**：对接人、理解需求、看图、拆任务、调度
- **子代理（Pro）= 程序员**：接任务、干活、交付结果
- 流程：用户(微信) → Omni(分析) → spawn Pro(执行) → 结果回传 → 微信通知
- **核心要求：主会话模型必须支持 tool calling**（spawn/exec/message），否则调度逻辑崩
- 子代理模型可随时替换，不绑死
- Omni 同时满足：tool calling + 看图，是最优项目经理人选
- Pro 推理能力强，适合干复杂活

## 子代理思考等级
- **写代码/架构设计** → `thinking: high`
- **读代码/分析** → `thinking: medium`（默认够用）
- **简单任务**（列表、查找）→ `thinking: low`

## 模型能力备注
- **MiMo-V2-Omni** 是唯一支持图片识别的模型（Anthropic兼容API能传图）
- **MiniMax 全家桶** 的 Anthropic 兼容 API 端点不传图片，只能当纯文本模型用
- 给不支持图片的模型在配置里加 `input: ["image"]` 无效，API不传图就是不能用
- **MiniMax 子代理的 provider 名字是 `minimax`**，不是 `minimax-cn`！config 里写的是 `minimax`，别自作聪明加 `-cn`
- MiniMax API key 已在 openclaw.json 的 `minimax` provider 里配置好，不需要额外的 auth-profiles.json

## GD32F527 技术经验
- **USB 48MHz 时钟**：200MHz 主频（25M HXTAL, PLL_N=400, PLL_Q=9）下 PLLQ=44.44MHz 不是精确48MHz。必须用 IRC48M 片内RC振荡器。
- **IRC48M 使用要点**：先 `rcu_osci_on(RCU_IRC48M)` 使能，等 `RCU_FLAG_IRC48MSTB` 稳定，再 `rcu_ck48m_clock_config(RCU_CK48MSRC_IRC48M)` 选源。不能只选源不使能，否则 USB reset 卡死。
- **PLL48M 时钟源选择**：`rcu_pll48m_clock_config()` 只有 PLLQ/PLLSAIP 两个选项，IRC48M 是在 CK48M mux 级别选的。
