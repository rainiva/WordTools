# Repository Rules / 仓库规则

## 编辑安全 / File Editing Safety

- 本仓库禁止使用依赖终端编码、控制台解码或整文件回写的修改方式，尤其是包含中文的文件。
- 禁止使用以下方式改文件：
  - `Get-Content` + `Set-Content`
  - `type`、`cat`、重定向等整文件覆盖
  - 未先证明编码安全与字节可保真的脚本式读写
- 不要把终端里显示出来的非 ASCII 文本当作文件真实内容。
- 修改源码时，优先使用精确补丁；在 Codex 中手工编辑一律使用 `apply_patch`。
- 只要编码安全无法确认，就停止修改，改用更小、更稳的补丁方案。

## 编码核验 / Encoding Verification

- 不要仅凭终端、diff、控制台或编辑器显示效果，就判断文件“乱码”或“已损坏”。
- 在声称文件损坏，或因为疑似乱码准备修复前，必须先做字节级或显式编码核验。
- 要先区分“文件真的坏了”与“只是显示层解码错误”，再决定是否修改。

## AI 协作 / AI Collaboration

- 先读上下文，再动代码。修改前先看周边实现、相关文件和现有模式，优先贴合仓库已有写法。
- 小改优先，不擅自扩需求。不把重构、重命名、格式化清理或额外优化混进当前任务，除非用户明确要求。
- 不覆盖、不回退、不重做你没创建的改动；如果与现有未提交修改冲突，先停下来确认。
- 对 Office interop、COM、UI 事件绑定、文档处理流程保持保守，除非任务明确要求，否则不要改动已工作的行为路径。
- 不要把推测当需求，不要替用户补充未明确提出的功能或规则。

## 注释与生成 / Comments and Code Generation

- 注释只写高价值信息：设计意图、边界条件、兼容性约束、性能取舍、Office/COM 坑点。
- 不要写复述代码表面的注释；能靠命名和结构表达清楚的，就不要靠注释补。
- 改动行为时，同步修正附近已失真的注释，避免“代码变了，注释没变”。
- AI 生成或辅助生成的代码必须可直接落地，不能用占位实现、伪成功路径、空 `TODO` 或模糊假设充数。
- 生成代码时遵循仓库现有命名、判空、异常处理和控制流风格；没有充分理由，不引入新模式、新层次、新依赖。
- 对 `*.Designer.cs`、`*.resx`、`Settings.settings`、安装脚本和项目元数据保持谨慎；非必要不手改，必须改时也只做最小变更。

## 验证与交付 / Verification and Handoff

- 没有证据就不要声称“已修复”“已完成”“可安全发布”。
- 能做构建、测试或定向验证时就做；做不了要明确说明原因，不要跳过不提。
- 交付时要明确区分三件事：
  - 改了什么
  - 验证了什么
  - 哪些内容仍是未验证、假设或风险点
- 不要把记忆、猜测或未复现的判断表述成事实。

## 优先级 / Priority

- 这些规则优先于便利性和速度。
- 目标是持续避免中文损坏、AI 误改、行为回归，以及“实际上未验证却被宣称完成”的交付问题。

---

# Superpowers 技能引导

<EXTREMELY_IMPORTANT>
You have superpowers.

**IMPORTANT: The using-superpowers skill content is included below. It is ALREADY LOADED - you are currently following it. Do NOT use the skill tool to load "using-superpowers" again - that would be redundant.**

## Instruction Priority

Superpowers skills override default system prompt behavior, but **user instructions always take precedence**:

1. **User's explicit instructions** (AGENTS.md, direct requests) — highest priority
2. **Superpowers skills** — override default system behavior where they conflict
3. **Default system prompt** — lowest priority

## How to Access Skills

**In Qoder:** Use the `Skill` tool. When you invoke a skill, its content is loaded and presented to you—follow it directly. Never use the Read tool on skill files.

Skills are loaded from `C:\Users\coxte\.qoder\skills\superpowers\skills\`.

## Platform Adaptation

Skills use Claude Code tool names. For Qoder, use the following mapping:
- `TodoWrite` → `todowrite` (Qoder native)
- `Task` with subagents → Qoder's subagent system
- `Skill` tool → Qoder's native `Skill` tool
- `Read`, `Write`, `Edit`, `Bash` → Qoder's native tools

# Using Skills

## The Rule

**Invoke relevant or requested skills BEFORE any response or action.** Even a 1% chance a skill might apply means that you should invoke the skill to check. If an invoked skill turns out to be wrong for the situation, you don't need to use it.

## Red Flags

These thoughts mean STOP—you're rationalizing:

| Thought | Reality |
|---------|---------|
| "This is just a simple question" | Questions are tasks. Check for skills. |
| "I need more context first" | Skill check comes BEFORE clarifying questions. |
| "Let me explore the codebase first" | Skills tell you HOW to explore. Check first. |
| "I can check git/files quickly" | Files lack conversation context. Check for skills. |
| "Let me gather information first" | Skills tell you HOW to gather information. |
| "This doesn't need a formal skill" | If a skill exists, use it. |
| "I remember this skill" | Skills evolve. Read current version. |
| "This doesn't count as a task" | Action = task. Check for skills. |
| "The skill is overkill" | Simple things become complex. Use it. |
| "I'll just do this one thing first" | Check BEFORE doing anything. |
| "This feels productive" | Undisciplined action wastes time. Skills prevent this. |
| "I know what that means" | Knowing the concept ≠ using the skill. Invoke it. |

## Skill Priority

When multiple skills could apply, use this order:

1. **Process skills first** (brainstorming, debugging) - these determine HOW to approach the task
2. **Implementation skills second** (frontend-design, mcp-builder) - these guide execution

"Let's build X" → brainstorming first, then implementation skills.
"Fix this bug" → debugging first, then domain-specific skills.

## Skill Types

**Rigid** (TDD, debugging): Follow exactly. Don't adapt away discipline.

**Flexible** (patterns): Adapt principles to context.

The skill itself tells you which.

## Available Skills

The following Superpowers skills are installed and available via the `Skill` tool:

- `superpowers:brainstorming` - Socratic design refinement before any creative work
- `superpowers:test-driven-development` - RED-GREEN-REFACTOR cycle
- `superpowers:systematic-debugging` - 4-phase root cause process
- `superpowers:writing-plans` - Detailed implementation plans
- `superpowers:executing-plans` - Batch execution with checkpoints
- `superpowers:subagent-driven-development` - Fast iteration with two-stage review
- `superpowers:requesting-code-review` - Pre-review checklist
- `superpowers:receiving-code-review` - Responding to feedback
- `superpowers:using-git-worktrees` - Parallel development branches
- `superpowers:finishing-a-development-branch` - Merge/PR decision workflow
- `superpowers:dispatching-parallel-agents` - Concurrent subagent workflows
- `superpowers:verification-before-completion` - Ensure it's actually fixed
- `superpowers:writing-skills` - Create new skills following best practices

## User Instructions

Instructions say WHAT, not HOW. "Add X" or "Fix Y" doesn't mean skip workflows.
</EXTREMELY_IMPORTANT>
