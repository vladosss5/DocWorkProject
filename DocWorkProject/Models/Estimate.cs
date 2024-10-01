namespace DocWorkProject.Models;

public class Estimate
{
    /// <summary>
    /// Id.
    /// </summary>
    public string Id { get; set; } = null!;

    /// <summary>
    /// Наименование.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// Первоначальная стоимость.
    /// </summary>
    public double InitialCost { get; set; }

    /// <summary>
    /// Остаточная стоимость.
    /// </summary>
    public double ResidualCost { get; set; }

    /// <summary>
    /// Родительской модели.
    /// </summary>
    public string? ParentModelId { get; set; }

    /// <summary>
    /// Родительская модель.
    /// </summary>
    public WordDataModel? ParentModel { get; set; }
}