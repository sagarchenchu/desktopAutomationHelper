using System;
using System.ComponentModel;
using System.Linq;

namespace DesktopRuntime.Fixtures;

public sealed record CustomerRecord(int Id, string Name, int Age, string Role);

public sealed class RecordStore
{
    private int nextId = 1;
    public BindingList<CustomerRecord> Records { get; } = new();
    public static string[] Roles { get; } = ["Reader", "Editor", "Administrator"];

    private static CustomerRecord Validate(int id, string name, string age, string role)
    {
        name = name.Trim();
        if (name.Length is < 1 or > 80) throw new ArgumentException("Name must contain 1–80 characters.");
        if (!int.TryParse(age, out int years) || years is < 1 or > 120)
            throw new ArgumentException("Age must be an integer from 1 to 120.");
        if (!Roles.Contains(role)) throw new ArgumentException("Select a role.");
        return new CustomerRecord(id, name, years, role);
    }

    public string Add(string name, string age, string role)
    {
        var record = Validate(nextId, name, age, role);
        Records.Add(record);
        nextId++;
        return $"Added record {record.Id}.";
    }
    public string Update(int id, string name, string age, string role)
    {
        var replacement = Validate(id, name, age, role);
        int index = Records.ToList().FindIndex(row => row.Id == id);
        if (index < 0) throw new ArgumentException("Select an existing row.");
        Records[index] = replacement;
        return $"Updated record {id}.";
    }
    public string Delete(int id)
    {
        var record = Records.FirstOrDefault(row => row.Id == id)
            ?? throw new ArgumentException("Select an existing row.");
        Records.Remove(record);
        return $"Deleted record {id}.";
    }
}
