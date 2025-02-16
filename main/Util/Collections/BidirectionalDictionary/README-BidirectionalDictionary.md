[![NuGet](https://img.shields.io/nuget/v/BidirectionalDictionary)](https://www.nuget.org/packages/BidirectionalDictionary/)
[![Downloads](https://img.shields.io/nuget/dt/BidirectionalDictionary)](https://www.nuget.org/packages/BidirectionalDictionary/)
[![License](https://img.shields.io/github/license/iiKuzmychov/BidirectionalDictionary)](https://github.com/iiKuzmychov/BidirectionalDictionary/blob/master/LICENSE.md)

# BidirectionalDictionary

Proper implementation of a bidirectional dictionary, also known as "BiMap" or "Two-way dictionary", for [.NET Standard 2.0](https://learn.microsoft.com/dotnet/standard/net-standard?tabs=net-standard-2-0#select-net-standard-version).

## Example of usage

```cs
using System.Collections.Generic;

var countryCapitalDictionary = new BidirectionalDictionary<string, string>()
{
    ["Italy"] = "Rome",
    ["India"] = "New Delhi",
    ["USA"]   = "Washington, D.C.",
};

Console.WriteLine(countryCapitalDictionary["Italy"]); // "Rome"
Console.WriteLine(countryCapitalDictionary.Inverse["Rome"]); // "Italy"
```

## Read-only support

If you need an read-only version of the bidirectional dictionary, the library provides
an easy way to achieve this.

Use for `BidirectionalDictionary`:

```cs
var readOnlyDictionary = dictionary.AsReadOnly();
```

Use for `IBidirectionalDictionary`:

```cs
using System.Collections.ObjectModel;

var readOnlyDictionary = new ReadOnlyBidirectionalDictionary<string, string>(dictionary);
```

## Interfaces

This library provides the following interfaces for greater flexibility:

- `IBidirectionalDictionary`
- `IReadOnlyBidirectionalDictionary`

Both `BidirectionalDictionary` and `ReadOnlyBidirectionalDictionary` implement these
interfaces, making it easier for you to work with various levels of abstractions.

## License

The library is licensed under the
[MIT](https://github.com/iiKuzmychov/BidirectionalDictionary/blob/master/LICENSE.md)
license.