namespace DocWorkProject.Models;

public class Document
{
    /// <summary>
    /// Название файла.
    /// </summary>
    public string Name { get; set; } = null!;

    /// <summary>
    /// Путь.
    /// </summary>
    public string Path { get; set; } = null!;
}