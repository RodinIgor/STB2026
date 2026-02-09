# STB2026 AI Connector for Revit

Open-source MCP сервер — подключает Claude Desktop / Cursor / Claude Code к Autodesk Revit.  
Альтернатива Nonica AI Connector с встроенными HVAC проверками по СП 60.13330.2020.

---

## Архитектура

```
┌─────────────────────────────────────┐
│  Claude Desktop / Cursor            │  AI-приложение
│  Connectors → Revit_STB2026 ✅      │
└────────────┬────────────────────────┘
             │ stdin/stdout (MCP JSON-RPC 2.0)
             ▼
┌─────────────────────────────────────┐
│  STB2026.McpServer.exe  (.NET 8)   │  Процесс-посредник
│  5 MCP tools для Claude            │
└────────────┬────────────────────────┘
             │ Named Pipe (length-prefixed JSON)
             ▼
┌─────────────────────────────────────┐
│  STB2026.RevitBridge.dll (Revit)   │  Плагин Revit
│  EventBridge → UI thread → API     │
│  Кнопка "A.I. Connector" на ribbon │
│  Auto/Manual настройка + тоглы     │
└─────────────────────────────────────┘
```

## 5 MCP Tools

| # | Tool | Что делает |
|---|------|-----------|
| 1 | `get_model_info` | Обзор проекта: категории, системы, уровни, виды |
| 2 | `get_elements` | Поиск элементов: by_category, by_view, by_selection, by_ids, by_system |
| 3 | `get_element_params` | Все параметры элементов (значения, единицы, id параметра) |
| 4 | `modify_model` | 9 действий: set_param, set_color, reset_color, select, move, rotate, delete, isolate, create_tag |
| 5 | `run_hvac_check` | 4 HVAC проверки: velocity (СП 60.13330), system_validation, wall_intersections, tag_ducts |

## Установка

### Требования
- Windows 10/11
- Autodesk Revit 2025
- .NET 8.0 Runtime (для McpServer)
- Claude Desktop / Cursor / Claude Code

### Шаг 1 — Сборка

```bash
dotnet build STB2026.MCP.sln -c Release
```

### Шаг 2 — Развёртывание

Скопируйте файлы в `C:\STB2026\`:

```
C:\STB2026\
├── STB2026.McpServer.exe          ← из STB2026.McpServer\bin\Release\net8.0-windows\
├── STB2026.McpServer.dll
├── STB2026.McpServer.deps.json
├── STB2026.McpServer.runtimeconfig.json
├── STB2026.Shared.dll
├── STB2026.RevitBridge.dll        ← из STB2026.RevitBridge\bin\Release\net48\
└── Newtonsoft.Json.dll
```

### Шаг 3 — Регистрация в Revit

Скопируйте `config\STB2026.Bridge.addin` в:
```
%APPDATA%\Autodesk\Revit\Addins\2025\
```

### Шаг 4 — Подключение Claude Desktop

**Вариант A — Auto (через UI Revit):**
1. Откройте Revit → вкладка STB2026 → кнопка "A.I. Connector"
2. Вкладка Auto → нажмите "Подключить" рядом с Claude Desktop
3. Перезапустите Claude Desktop

**Вариант B — Manual:**
1. Скопируйте содержимое `config\claude_desktop_config.json`
2. Вставьте в `%APPDATA%\Claude\claude_desktop_config.json`
3. Перезапустите Claude Desktop

### Шаг 5 — Проверка

В Claude Desktop напишите:
```
Покажи обзор проекта — какие системы вентиляции есть?
```

Claude вызовет `get_model_info` и покажет информацию о проекте.

## UI Настройки (кнопка "A.I. Connector")

### Вкладка Auto
- Автоматическое обнаружение Claude Desktop, Cursor, Claude Code
- Статусы: ✓ подключён / ○ требует настройки / ✗ не найден
- Кнопка "Подключить" — автоматическая правка конфига
- Статус Named Pipe сервера

### Вкладка Manual
- JSON для копирования в конфиг AI-приложения
- Кнопка Copy

### Тоглы безопасности
- **Enable Connection** — вкл/выкл всего MCP подключения
- **Enable A.I. Tools to Edit** — блокирует modify_model (read-only режим)

## Безопасность
- Named Pipe ACL: доступ только текущему пользователю Windows
- Тогл Edit Tools блокирует write-операции на уровне CommandRouter
- Бэкап claude_desktop_config.json перед изменением
- Все транзакции именованы "STB2026: ..." для отслеживания

## Структура проекта

```
STB2026.Shared/                          .NET Standard 2.0
└── Protocol/BridgeProtocol.cs           Request/Response + PipeIO framing

STB2026.McpServer/                       .NET 8.0-windows (console)
├── Program.cs                           MCP stdio transport
├── Pipe/RevitPipeClient.cs              Named Pipe клиент + auto-discovery
└── Tools/RevitTools.cs                  5 MCP tools

STB2026.RevitBridge/                     .NET Framework 4.8 + WPF (Revit 2025)
├── BridgeApp.cs                         IExternalApplication (entry point)
├── McpSettingsCommand.cs                IExternalCommand (кнопка ribbon)
├── Infrastructure/
│   ├── EventBridge.cs                   ExternalEvent → UI thread
│   ├── PipeServer.cs                    Named Pipe сервер + discovery
│   ├── CommandRouter.cs                 Маршрутизация + тогл Edit
│   └── McpSettingsManager.cs            Auto-detect AI apps + config management
├── Handlers/
│   ├── ParamHelper.cs                   Утилиты параметров (Revit 2025 API)
│   ├── ModelInfoHandler.cs              get_model_info
│   ├── ElementsHandler.cs              get_elements (5 режимов)
│   ├── ElementParamsHandler.cs          get_element_params
│   ├── ModifyHandler.cs                 modify_model (9 действий)
│   └── HvacCheckHandler.cs             run_hvac_check (4 проверки)
└── UI/
    ├── McpSettingsWindow.xaml            WPF окно настроек
    └── McpSettingsWindow.xaml.cs         Code-behind

config/
├── STB2026.Bridge.addin                 Manifest для Revit
└── claude_desktop_config.json           Пример конфига Claude Desktop
```

## Примеры использования

```
"Покажи обзор проекта"
→ get_model_info

"Найди все воздуховоды системы П1"
→ get_elements(mode="by_system", value="П1")

"Какой расход у воздуховода 123456?"
→ get_element_params(element_ids="123456")

"Проверь скорости воздуха по СП 60"
→ run_hvac_check(check_type="velocity")

"Окрась проблемные воздуховоды в красный"
→ modify_model(action="set_color", element_ids="...", data='{"r":255,"g":0,"b":0}')

"Найди пересечения воздуховодов со стенами"
→ run_hvac_check(check_type="wall_intersections")
```

## Исправленные баги (v2.0)

- ✅ JSON сериализация: McpServer и RevitBridge теперь используют совместимый PascalCase
- ✅ Revit 2025 API: удалены deprecated `ParameterType.YesNo` и `ParameterGroup`
- ✅ WPF окно настроек: Auto/Manual табы, тоглы, статусы

## Лицензия

Open-source. Свободное использование.
