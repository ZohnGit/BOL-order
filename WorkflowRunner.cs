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
        DateOnly targetDate,
        IProgress<WorkflowProgress> progress,
        CancellationToken cancellationToken = default)
    {
        _lastDetailToken = null;
        var target = targetDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        var activeOrders = new List<OrderSummary>();
        var cancelledOrders = new List<CancelledOrder>();
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
                    cancelledOrders.Add(new CancelledOrder(
                        order.OrderId ?? string.Empty,
                        order.OrderPlacedDateTime ?? string.Empty));
                }
                else
                {
                    activeOrders.Add(order);
                }
            }

            if (orders.Count == 0) break;

            var lastDate = DatePart(orders[^1].OrderPlacedDateTime);
            var shouldContinue = lastDate is not null && string.CompareOrdinal(lastDate, target) >= 0;
            if (!shouldContinue) break;
            page++;
        }

        // Mirrors “展开目标日期订单”: only the normal-order path is filtered here.
        var selectedOrders = activeOrders
            .Where(order => DatePart(order.OrderPlacedDateTime) == target)
            .ToList();

        progress.Report(new WorkflowProgress(20, $"已找到 {selectedOrders.Count} 个待处理订单。"));

        var rows = new List<ExportRow>();
        for (var index = 0; index < selectedOrders.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var order = selectedOrders[index];
            if (string.IsNullOrWhiteSpace(order.OrderId)) continue;

            // Mirrors the Wait node before each item.
            await Task.Delay(TimeSpan.FromSeconds(1.5), cancellationToken);

            // Mirrors orderCounter === 1 || (orderCounter - 1) % 15 === 0.
            var orderNumber = index + 1;
            string detailToken;
            if (orderNumber == 1 || (orderNumber - 1) % 15 == 0)
            {
                progress.Report(ItemProgress(index, selectedOrders.Count,
                    $"正在刷新登录信息（第 {orderNumber}/{selectedOrders.Count} 单）…"));
                detailToken = await _api.LoginAsync(user, password, cancellationToken);
            }
            else
            {
                // The n8n workflow reuses the last token. Keep it explicitly here.
                detailToken = _lastDetailToken ?? initialToken;
            }
            _lastDetailToken = detailToken;

            progress.Report(ItemProgress(index, selectedOrders.Count,
                $"正在处理第 {orderNumber}/{selectedOrders.Count} 单：{order.OrderId}"));

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

            progress.Report(ItemProgress(index + 1, selectedOrders.Count,
                $"已处理 {orderNumber}/{selectedOrders.Count} 单。"));
        }

        progress.Report(new WorkflowProgress(95, "正在生成 XLSX 文件…"));
        var xlsx = XlsxWriter.Create(rows);
        progress.Report(new WorkflowProgress(100, "处理完成，XLSX 文件已可以获取。"));
        return new WorkflowResult(cancelledOrders, rows, xlsx);
    }

    private string? _lastDetailToken;

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
}
