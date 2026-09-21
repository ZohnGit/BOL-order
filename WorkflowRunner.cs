using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace BolOrderExporter;

public sealed record WorkflowProgress(double Percentage, string Status, bool IsIndeterminate = false);

public sealed class WorkflowRunner
{
    private readonly BolApiClient _api;

    public WorkflowRunner(BolApiClient api) => _api = api;

    public async Task<WorkflowResult> RunAsync(
        string user,
        string password,
        string initialToken,
        DateOnly startDate,
        DateOnly endDate,
        IProgress<WorkflowProgress> progress,
        CancellationToken cancellationToken = default)
    {
        _lastDetailToken = null;
        var start = startDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        var end = endDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        var activeOrders = new List<OrderSummary>();
        var cancelledOrderSummaries = new List<OrderSummary>();
        var page = 1;

        progress.Report(new WorkflowProgress(5, "正在分页获取订单…", true));

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress.Report(new WorkflowProgress(5, $"正在获取第 {page} 页订单…", true));

            using var pageDocument = await _api.GetOrdersAsync(initialToken, page, cancellationToken);
            var orders = ReadOrders(pageDocument.RootElement);

            foreach (var order in orders)
            {
                if (order.IsCancelOnlyOrder)
                {
                    cancelledOrderSummaries.Add(order);
                }
                else
                {
                    activeOrders.Add(order);
                }
            }

            if (orders.Count == 0) break;

            var lastDate = DatePart(orders[^1].OrderPlacedDateTime);
            var shouldContinue = lastDate is not null && string.CompareOrdinal(lastDate, start) >= 0;
            if (!shouldContinue) break;
            page++;
        }

        var selectedOrders = activeOrders
            .Where(order => DateInRange(DatePart(order.OrderPlacedDateTime), start, end))
            .ToList();
        var selectedCancelledOrders = cancelledOrderSummaries
            .Where(order => DateInRange(DatePart(order.OrderPlacedDateTime), start, end))
            .ToList();

        var totalDetailOrders = selectedOrders.Count + selectedCancelledOrders.Count;
        progress.Report(new WorkflowProgress(20, $"已找到 {selectedOrders.Count} 个正常订单，{selectedCancelledOrders.Count} 个取消单。"));

        var rows = new List<ExportRow>();
        var cancelledOrders = new List<CancelledOrder>();
        var processedDetailOrders = 0;

        for (var index = 0; index < selectedOrders.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var order = selectedOrders[index];
            if (string.IsNullOrWhiteSpace(order.OrderId)) continue;

            var orderNumber = processedDetailOrders + 1;
            var detailToken = await GetDetailTokenAsync(user, password, initialToken, orderNumber, totalDetailOrders, progress, cancellationToken);

            progress.Report(ItemProgress(processedDetailOrders, totalDetailOrders,
                $"正在处理正常订单 {orderNumber}/{totalDetailOrders}：{order.OrderId}"));

            using var orderDocument = await _api.GetOrderAsync(detailToken, order.OrderId, cancellationToken);
            using var shipmentsDocument = await _api.GetShipmentsAsync(detailToken, order.OrderId, cancellationToken);
            var shipmentId = FirstShipmentId(shipmentsDocument.RootElement);
            if (string.IsNullOrWhiteSpace(shipmentId))
            {
                // Keep the order in the XLSX. Only shipment-derived fields stay empty.
                rows.AddRange(BuildExportRows(orderDocument.RootElement, default));
            }
            else
            {
                using var shipmentDocument = await _api.GetShipmentAsync(detailToken, shipmentId, cancellationToken);
                rows.AddRange(BuildExportRows(orderDocument.RootElement, shipmentDocument.RootElement));
            }

            processedDetailOrders++;
            progress.Report(ItemProgress(processedDetailOrders, totalDetailOrders,
                $"已处理 {processedDetailOrders}/{totalDetailOrders} 单。"));
        }

        for (var index = 0; index < selectedCancelledOrders.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var order = selectedCancelledOrders[index];
            if (string.IsNullOrWhiteSpace(order.OrderId)) continue;

            var orderNumber = processedDetailOrders + 1;
            var detailToken = await GetDetailTokenAsync(user, password, initialToken, orderNumber, totalDetailOrders, progress, cancellationToken);

            progress.Report(ItemProgress(processedDetailOrders, totalDetailOrders,
                $"正在处理取消单 {orderNumber}/{totalDetailOrders}：{order.OrderId}"));

            using var orderDocument = await _api.GetOrderAsync(detailToken, order.OrderId, cancellationToken);
            var cancelledRows = BuildCancelledOrders(orderDocument.RootElement, order.OrderId, order.OrderPlacedDateTime).ToList();
            if (cancelledRows.Count == 0)
            {
                cancelledOrders.Add(new CancelledOrder(order.OrderId, order.OrderPlacedDateTime ?? string.Empty, null));
            }
            else
            {
                cancelledOrders.AddRange(cancelledRows);
            }

            processedDetailOrders++;
            progress.Report(ItemProgress(processedDetailOrders, totalDetailOrders,
                $"已处理 {processedDetailOrders}/{totalDetailOrders} 单。"));
        }

        progress.Report(new WorkflowProgress(95, "正在生成 XLSX 文件…"));
        var xlsx = XlsxWriter.Create(rows);
        progress.Report(new WorkflowProgress(100, "处理完成，XLSX 文件已可以获取。"));
        return new WorkflowResult(cancelledOrders, rows, xlsx);
    }

    private string? _lastDetailToken;

    private async Task<string> GetDetailTokenAsync(
        string user,
        string password,
        string initialToken,
        int orderNumber,
        int total,
        IProgress<WorkflowProgress> progress,
        CancellationToken cancellationToken)
    {
        // Mirrors the Wait node before each item.
        await Task.Delay(TimeSpan.FromSeconds(1.5), cancellationToken);

        // Mirrors orderCounter === 1 || (orderCounter - 1) % 15 === 0.
        if (orderNumber == 1 || (orderNumber - 1) % 15 == 0)
        {
            progress.Report(ItemProgress(orderNumber - 1, total,
                $"正在刷新登录信息（第 {orderNumber}/{total} 单）…"));
            _lastDetailToken = await _api.LoginAsync(user, password, cancellationToken);
        }
        else
        {
            _lastDetailToken ??= initialToken;
        }

        return _lastDetailToken;
    }

    private static WorkflowProgress ItemProgress(int completed, int total, string status)
    {
        var percentage = total == 0 ? 90 : 20 + 70d * completed / total;
        return new WorkflowProgress(percentage, status);
    }

    private static List<OrderSummary> ReadOrders(JsonElement root)
    {
        var result = new List<OrderSummary>();
        if (!TryArray(root, "orders", out var orders)) return result;

        foreach (var order in orders.EnumerateArray())
        {
            var itemsExist = TryArray(order, "orderItems", out var items);
            var totalItems = itemsExist ? items.GetArrayLength() : 0;
            var activeItems = 0;
            if (itemsExist)
            {
                foreach (var item in items.EnumerateArray())
                {
                    var hasQ = TryDecimal(item, "quantity", out var quantity);
                    var hasCancelled = TryDecimal(item, "quantityCancelled", out var cancelled);
                    if (!hasQ || !hasCancelled || quantity - cancelled > 0) activeItems++;
                }
            }

            BolApiClient.TryString(order, "orderId", out var orderId);
            BolApiClient.TryString(order, "orderPlacedDateTime", out var placed);
            result.Add(new OrderSummary(orderId, placed, totalItems > 0 && activeItems == 0));
        }
        return result;
    }

    private static string? FirstShipmentId(JsonElement root)
    {
        if (!TryArray(root, "shipments", out var shipments) || shipments.GetArrayLength() == 0) return null;
        var first = shipments[0];
        return BolApiClient.TryString(first, "shipmentId", out var id) ? id : null;
    }

    private static IEnumerable<ExportRow> BuildExportRows(JsonElement order, JsonElement shipment)
    {
        BolApiClient.TryString(order, "orderId", out var orderId);
        BolApiClient.TryString(order, "orderPlacedDateTime", out var paymentTime);

        var shipmentDetails = Child(order, "shipmentDetails");
        var billingDetails = Child(order, "billingDetails");
        BolApiClient.TryString(shipmentDetails, "firstName", out var firstName);
        BolApiClient.TryString(shipmentDetails, "surname", out var surname);
        var buyerName = string.Join(" ", new[] { firstName, surname }.Where(x => !string.IsNullOrWhiteSpace(x))).Trim();
        if (buyerName.Length == 0) buyerName = null;

        BolApiClient.TryString(shipmentDetails, "email", out var buyerEmail);
        if (string.IsNullOrWhiteSpace(buyerEmail)) BolApiClient.TryString(billingDetails, "email", out buyerEmail);
        BolApiClient.TryString(shipmentDetails, "countryCode", out var countryCode);

        var transport = Child(shipment, "transport");
        BolApiClient.TryString(transport, "trackAndTrace", out var trackAndTrace);

        if (!TryArray(order, "orderItems", out var items) || items.GetArrayLength() == 0) yield break;
        foreach (var item in items.EnumerateArray())
        {
            TryDecimal(item, "quantity", out var quantity);
            TryDecimal(item, "quantityCancelled", out var cancelled);
            var net = quantity - cancelled;
            if (net <= 0) continue;

            decimal? amount = TryDecimal(item, "totalPrice", out var totalPrice) ? totalPrice : null;
            var offer = Child(item, "offer");
            var product = Child(item, "product");
            BolApiClient.TryString(offer, "reference", out var sku);
            if (string.IsNullOrWhiteSpace(sku)) BolApiClient.TryString(product, "ean", out sku);
            BolApiClient.TryString(item, "latestChangedDateTime", out var changed);

            yield return new ExportRow(
                orderId, buyerName, buyerEmail, paymentTime, countryCode, trackAndTrace,
                amount, sku, net, quantity, cancelled, changed);
        }
    }

    private static IEnumerable<CancelledOrder> BuildCancelledOrders(JsonElement order, string fallbackOrderId, string? fallbackPlaced)
    {
        BolApiClient.TryString(order, "orderId", out var orderId);
        if (string.IsNullOrWhiteSpace(orderId)) orderId = fallbackOrderId;

        BolApiClient.TryString(order, "orderPlacedDateTime", out var placed);
        if (string.IsNullOrWhiteSpace(placed)) placed = fallbackPlaced ?? string.Empty;

        if (!TryArray(order, "orderItems", out var items) || items.GetArrayLength() == 0)
        {
            yield return new CancelledOrder(orderId ?? string.Empty, placed ?? string.Empty, null);
            yield break;
        }

        foreach (var item in items.EnumerateArray())
        {
            TryDecimal(item, "quantityCancelled", out var cancelled);
            if (cancelled <= 0) continue;

            var offer = Child(item, "offer");
            var product = Child(item, "product");
            BolApiClient.TryString(offer, "reference", out var sku);
            if (string.IsNullOrWhiteSpace(sku)) BolApiClient.TryString(product, "ean", out sku);

            yield return new CancelledOrder(orderId ?? string.Empty, placed ?? string.Empty, sku);
        }
    }

    private static bool TryArray(JsonElement element, string name, out JsonElement array)
    {
        array = default;
        return element.ValueKind == JsonValueKind.Object &&
               element.TryGetProperty(name, out array) &&
               array.ValueKind == JsonValueKind.Array;
    }

    private static JsonElement Child(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var child)
            ? child
            : default;

    private static bool TryDecimal(JsonElement element, string name, out decimal value)
    {
        value = 0;
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var property)) return false;
        if (property.ValueKind == JsonValueKind.Number) return property.TryGetDecimal(out value);
        return property.ValueKind == JsonValueKind.String &&
               decimal.TryParse(property.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    private static string? DatePart(string? value) =>
        string.IsNullOrWhiteSpace(value) || value.Length < 10 ? null : value[..10];

    private static bool DateInRange(string? value, string start, string end) =>
        value is not null &&
        string.CompareOrdinal(value, start) >= 0 &&
        string.CompareOrdinal(value, end) <= 0;
}
