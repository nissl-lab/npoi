using BenchmarkDotNet.Running;
using NPOI.Benchmarks;

BenchmarkSwitcher.FromAssembly(typeof(LargeExcelFileBenchmark).Assembly).Run();