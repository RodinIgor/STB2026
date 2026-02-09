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
    /// 5 MCP tools — полный доступ Claude к модели Revit.
    /// 
    /// Имена tool'ов = имена методов (snake_case).
    /// SDK автоматически берёт имя метода как tool name.
    /// </summary>
    [McpServerToolType]
    public sealed class RevitTools
    {
        private readonly RevitPipeClient _pipe;

        public RevitTools(RevitPipeClient pipe)
        {
            _pipe = pipe;
        }

        // ═══════════════════════════════════════════════════════
        //  TOOL 1: get_model_info
        // ═══════════════════════════════════════════════════════

        [McpServerTool, Description(
            "Получить обзор открытого проекта Revit: имя проекта, активный вид, " +
            "список категорий с количеством элементов, MEP-системы, уровни, единицы. " +
            "Вызывай первым для понимания контекста.")]
        public async Task<string> get_model_info(CancellationToken ct)
        {
            return await CallRevitAsync("get_model_info", new Dictionary<string, object>(), ct);
        }

        // ═══════════════════════════════════════════════════════
        //  TOOL 2: get_elements
        // ═══════════════════════════════════════════════════════

        [McpServerTool, Description(
            "Запросить элементы из модели Revit. Возвращает список с ID, именем, типом. " +
            "Режимы (mode): by_category (напр. 'Ducts','Walls','Rooms'), " +
            "by_view (элементы на виде), by_selection (текущее выделение), " +
            "by_ids (конкретные ID), by_system (элементы MEP-системы), " +
            "by_linked (элементы из связанного файла, value='имя_связи:категория', " +
            "напр. 'ТестАР2026:Rooms' или ':Walls' для всех связей, " +
            "или просто имя связи без категории для сводки). " +
            "limit ограничивает количество (умолч. 500).")]
        public async Task<string> get_elements(
            [Description("Режим: 'by_category','by_view','by_selection','by_ids','by_system','by_linked'")] string mode,
            [Description("Значение: имя категории, ID вида, список ID, имя системы, 'имя_связи:категория'. Для by_selection — пусто.")] string value = "",
            [Description("Макс. кол-во элементов")] int limit = 500,
            CancellationToken ct = default)
        {
            var p = new Dictionary<string, object>
            {
                ["mode"] = mode,
                ["value"] = value,
                ["limit"] = limit
            };
            return await CallRevitAsync("get_elements", p, ct);
        }

        // ═══════════════════════════════════════════════════════
        //  TOOL 3: get_element_params
        // ═══════════════════════════════════════════════════════

        [McpServerTool, Description(
            "Получить ВСЕ параметры для элементов (имя, значение, единицы, id параметра). " +
            "Один ID — все параметры. Несколько ID + param_name — одно значение для каждого.")]
        public async Task<string> get_element_params(
            [Description("ID элементов через запятую: '12345' или '12345,12346'")] string element_ids,
            [Description("Имя конкретного параметра (необязательно). Пусто = все параметры.")] string param_name = "",
            CancellationToken ct = default)
        {
            var p = new Dictionary<string, object>
            {
                ["element_ids"] = element_ids,
                ["param_name"] = param_name ?? ""
            };
            return await CallRevitAsync("get_element_params", p, ct);
        }

        // ═══════════════════════════════════════════════════════
        //  TOOL 4: modify_model
        // ═══════════════════════════════════════════════════════

        [McpServerTool, Description(
            "Модификация модели Revit. Действия (action): " +
            "'set_param' — установить параметр (data: param_name + param_value), " +
            "'set_color' — окрасить (data: r,g,b), " +
            "'reset_color' — сбросить цвета, " +
            "'select' — выделить в Revit, " +
            "'move' — переместить (data: dx,dy,dz в мм), " +
            "'rotate' — повернуть (data: angle в градусах), " +
            "'delete' — удалить (НЕОБРАТИМО!), " +
            "'isolate' — изолировать на виде, " +
            "'create_tag' — создать марку (data: tag_family).")]
        public async Task<string> modify_model(
            [Description("Действие: 'set_param','set_color','reset_color','select','move','rotate','delete','isolate','create_tag'")] string action,
            [Description("ID элементов через запятую")] string element_ids,
            [Description("JSON-данные. Примеры: set_color:{\"r\":255,\"g\":0,\"b\":0} | move:{\"dx\":100,\"dy\":0,\"dz\":0}")] string data = "{}",
            CancellationToken ct = default)
        {
            var p = new Dictionary<string, object>
            {
                ["action"] = action,
                ["element_ids"] = element_ids,
                ["data"] = data
            };
            return await CallRevitAsync("modify_model", p, ct);
        }

        // ═══════════════════════════════════════════════════════
        //  TOOL 5: run_hvac_check
        // ═══════════════════════════════════════════════════════

        [McpServerTool, Description(
            "Проверка систем ОВиК (HVAC). Типы (check_type): " +
            "'velocity' — скорости воздуха по СП 60.13330.2020, " +
            "'system_validation' — нулевой расход, без системы, отключённые, " +
            "'wall_intersections' — пересечения воздуховодов со стенами (координаты), " +
            "'tag_ducts' — автомаркировка воздуховодов. " +
            "system_name фильтрует по системе.")]
        public async Task<string> run_hvac_check(
            [Description("Тип: 'velocity','system_validation','wall_intersections','tag_ducts'")] string check_type,
            [Description("Имя системы для фильтра (пусто = все)")] string system_name = "",
            CancellationToken ct = default)
        {
            var p = new Dictionary<string, object>
            {
                ["check_type"] = check_type,
                ["system_name"] = system_name ?? ""
            };
            return await CallRevitAsync("run_hvac_check", p, ct);
        }

        // ═══════════════════════════════════════════════════════
        //  TOOL 6: manage_families
        // ═══════════════════════════════════════════════════════

        [McpServerTool, Description(
            "Управление семействами и общими параметрами Revit. Действия (action): " +
            "'list_families' — список семейств (data: {\"category\":\"...\"} для фильтра), " +
            "'load_family' — загрузить .rfa (data: {\"path\":\"C:\\\\...\",\"overwrite\":true}), " +
            "'list_shared_params' — список привязанных общих параметров, " +
            "'add_shared_param' — добавить общий параметр (data: {\"param_name\":\"...\",\"categories\":[\"Воздуховоды\"],\"is_instance\":true,\"group_name\":\"STB2026\",\"param_group\":\"Механизмы\"}), " +
            "'remove_shared_param' — удалить привязку (data: {\"param_name\":\"...\"}), " +
            "'list_family_types' — типоразмеры (data: {\"family_name\":\"...\"}), " +
            "'set_family_type_param' — изменить параметр типа (data: {\"family_name\":\"...\",\"type_name\":\"...\",\"param_name\":\"...\",\"param_value\":\"...\"}), " +
            "'duplicate_type' — дублировать типоразмер (data: {\"family_name\":\"...\",\"type_name\":\"...\",\"new_name\":\"...\"}), " +
            "'edit_family' — редактировать семейство (data: {\"family_name\":\"...\",\"params\":[{\"name\":\"...\",\"value\":\"...\",\"is_instance\":true}]}).")]
        public async Task<string> manage_families(
            [Description("Действие: 'list_families','load_family','list_shared_params','add_shared_param','remove_shared_param','list_family_types','set_family_type_param','duplicate_type','edit_family'")] string action,
            [Description("JSON-данные для действия")] string data = "{}",
            CancellationToken ct = default)
        {
            var p = new Dictionary<string, object>
            {
                ["action"] = action,
                ["data"] = data
            };
            return await CallRevitAsync("manage_families", p, ct);
        }

        // ═══════════════════════════════════════════════════════
        //  Pipe → Revit Bridge
        // ═══════════════════════════════════════════════════════

        private async Task<string> CallRevitAsync(string method, Dictionary<string, object> p, CancellationToken ct)
        {
            var request = new BridgeRequest
            {
                Method = method,
                Params = p
            };

            var response = await _pipe.SendAsync(request, ct);

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
