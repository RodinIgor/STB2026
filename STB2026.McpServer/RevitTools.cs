using System.Collections.Generic;
using System.ComponentModel;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using STB2026.McpServer.Pipe;
using STB2026.Shared.Protocol;

namespace STB2026.McpServer.Tools
{
    /// <summary>
    /// 11 MCP tools — полный доступ Claude к модели Revit.
    /// </summary>
    [McpServerToolType]
    public sealed class RevitTools
    {
        private readonly RevitPipeClient _pipe;
        public RevitTools(RevitPipeClient pipe) => _pipe = pipe;

        // ═══ TOOL 1: get_model_info ═══

        [McpServerTool, Description(
            "Получить обзор открытого проекта Revit: имя проекта, активный вид, " +
            "список категорий с количеством элементов, MEP-системы, уровни, единицы. " +
            "Вызывай первым для понимания контекста.")]
        public Task<string> get_model_info(CancellationToken ct)
            => CallAsync("get_model_info", new(), ct);

        // ═══ TOOL 2: get_elements ═══

        [McpServerTool, Description(
            "Запросить элементы из модели Revit. Возвращает список с ID, именем, типом. " +
            "Режимы (mode): by_category, by_view, by_selection, by_ids, by_system, by_linked. " +
            "limit ограничивает количество (умолч. 500).")]
        public Task<string> get_elements(
            [Description("Режим: 'by_category','by_view','by_selection','by_ids','by_system','by_linked'")] string mode,
            [Description("Значение: имя категории, ID вида, список ID, имя системы, 'имя_связи:категория'.")] string value = "",
            [Description("Макс. кол-во элементов")] int limit = 500,
            CancellationToken ct = default)
            => CallAsync("get_elements", new() { ["mode"] = mode, ["value"] = value, ["limit"] = limit }, ct);

        // ═══ TOOL 3: get_element_params ═══

        [McpServerTool, Description(
            "Получить ВСЕ параметры для элементов (имя, значение, единицы). " +
            "Один ID — все параметры. Несколько ID + param_name — одно значение для каждого.")]
        public Task<string> get_element_params(
            [Description("ID элементов через запятую")] string element_ids,
            [Description("Имя параметра (пусто = все)")] string param_name = "",
            CancellationToken ct = default)
            => CallAsync("get_element_params", new() { ["element_ids"] = element_ids, ["param_name"] = param_name ?? "" }, ct);

        // ═══ TOOL 4: modify_model ═══

        [McpServerTool, Description(
            "Модификация элементов Revit. Действия: " +
            "set_param, set_color, reset_color, select, move, rotate, delete, isolate, create_tag.")]
        public Task<string> modify_model(
            [Description("Действие")] string action,
            [Description("ID элементов через запятую")] string element_ids,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("modify_model", new() { ["action"] = action, ["element_ids"] = element_ids, ["data"] = data }, ct);

        // ═══ TOOL 5: run_hvac_check ═══

        [McpServerTool, Description(
            "Проверка систем ОВиК. Типы: velocity, system_validation, wall_intersections, tag_ducts.")]
        public Task<string> run_hvac_check(
            [Description("Тип проверки")] string check_type,
            [Description("Имя системы (пусто = все)")] string system_name = "",
            CancellationToken ct = default)
            => CallAsync("run_hvac_check", new() { ["check_type"] = check_type, ["system_name"] = system_name ?? "" }, ct);

        // ═══ TOOL 6: manage_families ═══

        [McpServerTool, Description(
            "Управление семействами и общими параметрами. Действия: " +
            "list_families, load_family, list_shared_params, add_shared_param, remove_shared_param, " +
            "list_family_types, set_family_type_param, duplicate_type, edit_family.")]
        public Task<string> manage_families(
            [Description("Действие")] string action,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("manage_families", new() { ["action"] = action, ["data"] = data }, ct);

        // ═══ TOOL 7: manage_sheets ═══

        [McpServerTool, Description(
            "Управление листами и видовыми экранами. Действия: " +
            "create_sheet, place_view, list_viewports, move_viewport, align_viewports.")]
        public Task<string> manage_sheets(
            [Description("Действие")] string action,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("manage_sheets", new() { ["action"] = action, ["data"] = data }, ct);

        // ═══ TOOL 8: manage_views ═══

        [McpServerTool, Description(
            "Управление видами Revit. Действия: " +
            "list_views (фильтр: type, prefix), create_plan, create_section, create_3d, " +
            "duplicate_view (mode: Duplicate/WithDetailing/AsDependent), " +
            "delete_view, rename_view, set_active, set_scale, " +
            "set_template, remove_template, list_templates, " +
            "set_crop_box, toggle_crop, set_detail_level (coarse/medium/fine), " +
            "set_display_style (wireframe/hidden/shaded/realistic), " +
            "get_view_filters, add_view_filter, remove_view_filter, " +
            "set_category_visibility (data: {\"view_id\",\"category\",\"visible\":bool}).")]
        public Task<string> manage_views(
            [Description("Действие")] string action,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("manage_views", new() { ["action"] = action, ["data"] = data }, ct);

        // ═══ TOOL 9: create_elements ═══

        [McpServerTool, Description(
            "Создание элементов в модели. Действия: " +
            "place_instance (family+type+xyz), create_wall, create_floor, create_ceiling, " +
            "create_room, create_space, create_duct (с размерами и системой), " +
            "create_pipe, create_flex_duct, create_insulation, " +
            "create_opening, create_text_note, create_dimension.")]
        public Task<string> create_elements(
            [Description("Действие")] string action,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("create_elements", new() { ["action"] = action, ["data"] = data }, ct);

        // ═══ TOOL 10: query_geometry ═══

        [McpServerTool, Description(
            "Геометрические запросы и измерения. Действия: " +
            "measure_distance, measure_length, get_bounding_box, get_location, " +
            "get_connectors (MEP коннекторы), get_room_at_point, get_room_boundaries, " +
            "find_nearest (ближайший элемент категории к точке), " +
            "check_intersections (пересечения BBox), get_area, get_volume.")]
        public Task<string> query_geometry(
            [Description("Действие")] string action,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("query_geometry", new() { ["action"] = action, ["data"] = data }, ct);

        // ═══ TOOL 11: manage_project ═══

        [McpServerTool, Description(
            "Управление проектом Revit. Действия: " +
            "get_project_info, set_project_info, list_phases, list_worksets, " +
            "list_design_options, list_links, reload_link, " +
            "list_warnings (предупреждения модели с группировкой), " +
            "get_units, list_materials, list_line_styles, list_fill_patterns, " +
            "purge_unused (анализ неиспользуемых типов, без удаления).")]
        public Task<string> manage_project(
            [Description("Действие")] string action,
            [Description("JSON-данные")] string data = "{}",
            CancellationToken ct = default)
            => CallAsync("manage_project", new() { ["action"] = action, ["data"] = data }, ct);

        // ═══ Pipe → Revit Bridge ═══

        private async Task<string> CallAsync(string method, Dictionary<string, object> p, CancellationToken ct)
        {
            var response = await _pipe.SendAsync(new BridgeRequest { Method = method, Params = p }, ct);
            if (!response.Success)
                return $"Ошибка: {response.Error}";
            return JsonSerializer.Serialize(response.Data, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
        }
    }
}
