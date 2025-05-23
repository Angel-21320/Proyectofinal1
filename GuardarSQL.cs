static void GuardarInvestigacion(string prompt, string resultado)
{
    try
    {
        var connectionString = "Server=LAPTOP-5NF567DI\\SQLEXPRESS;Database=Proyecto 1;Trusted_Connection=True;";
        using var conn = new SqlConnection(connectionString);
        conn.Open();
        var cmd = new SqlCommand("INSERT INTO Investigaciones (Prompt, Resultado) VALUES (@p, @r)", conn);
        cmd.Parameters.AddWithValue("@p", prompt);
        cmd.Parameters.AddWithValue("@r", resultado);
        cmd.ExecuteNonQuery();
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error al guardar en SQL: {ex.Message}");
    }
}