# Win32-Auto-Typer ⌨️

A system-level Windows automation tool written in PowerShell, designed to eliminate repetitive administrative tasks through direct thread manipulation.

## 🌟 Tech Highlights & C-Language Foundation
While working with hardware and low-level system operations, I realized that many internal efficiency bottlenecks could be solved by bridging high-level scripting with low-level system APIs.

Instead of relying on unstable external macro tools, this script directly invokes Windows Core APIs (`user32.dll`) to achieve precision:
- **Thread Attachment**: Utilized `AttachThreadInput` and `GetWindowThreadProcessId` (concepts translated from my C programming background) to bypass standard focus restrictions.
- **Direct Memory Injection**: Implemented `SendInput` and `SendMessage` to simulate keystrokes and GUI interactions directly at the OS message-queue level, ensuring 100% execution accuracy without mouse-takeover.
- **Dynamic UI**: Built a lightweight Windows Forms GUI on the fly without heavy compiled executables.

## 🚀 Use Case
Designed as an internal efficiency engineering tool to automate data entry, template generation, and repetitive GUI operations, significantly reducing manual overhead.
