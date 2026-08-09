using Xunit;

// CLI tests redirect Console.Out/Error; keep the assembly sequential to avoid races.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
