namespace BolOrderExporter;

public sealed record SavedCredential(string User, string Password);

public sealed record CancelledOrder(string OrderId, string OrderPlacedDateTime);

public sealed record OrderSummary(
    string? OrderId,
    string? OrderPlacedDateTime,
    bool IsCancelOnlyOrder);

public sealed record ExportRow(
    string? OrderId,
    string? BuyerName,
    string? BuyerEmail,
    string? PaymentTime,
    string? CountryCode,
    string? TrackAndTrace,
    decimal? Amount,
    string? Sku,
    decimal Quantity,
    decimal QuantityOriginal,
    decimal QuantityCancelled,
    string? ShipmentDateTime);

public sealed record WorkflowResult(
    IReadOnlyList<CancelledOrder> CancelledOrders,
    IReadOnlyList<ExportRow> ExportRows,
    byte[] XlsxBytes);
