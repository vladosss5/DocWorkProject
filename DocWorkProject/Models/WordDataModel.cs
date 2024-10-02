using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DocWorkProject.Models;

public class WordDataModel
{
    /// <summary>
    /// Id.
    /// </summary>
    public string Id { get; set; } = null!;

    /// <summary>
    /// Дата составления.
    /// </summary>
    public DateOnly DateCompilation { get; set; }

    /// <summary>
    /// Номер отчёта.
    /// </summary>
    public string? Number { get; set; }

    /// <summary>
    /// Заказчик.
    /// </summary>
    public string? Customer { get; set; }

    /// <summary>
    /// Оценщик.
    /// </summary>
    public string? Appraiser { get; set; }

    /// <summary>
    /// Тип стоимости.
    /// </summary>
    public string? TypeCost { get; set; }

    /// <summary>
    /// Цель оценки.
    /// </summary>
    public string? PurposeAssessment { get; set; }

    /// <summary>
    /// Дата оценки.
    /// </summary>
    public DateOnly DateAssessment { get; set; }

    /// <summary>
    /// Дата составления.
    /// </summary>
    public DateOnly DateCompilationReport { get; set; }

    /// <summary>
    /// Список оценок.
    /// </summary>
    public virtual List<Estimate> Estimates { get; set; } = new List<Estimate>();
}