[![NuGet](https://img.shields.io/nuget/dt/npoi)](https://www.nuget.org/packages/NPOI)
[![Nuget Ranking!](https://img.shields.io/badge/Nuget%20Ranking-282-red.svg)](https://github.com/nissl-lab/npoi/issues/1532)
[![Discord](https://img.shields.io/badge/Chat-Discord-d82679.svg?logo=discord&logoColor=white)](https://discord.gg/rABUseUHme)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg?style=flat-square&logo=Apache)](LICENSE)
[![EULA](https://img.shields.io/badge/EULA-OSMF-blue?labelColor=black&color=C9FF30)](https://github.com/nissl-lab/npoi/blob/master/OSMFEULA.txt)
[![traffic](https://api.segment.io/v1/pixel/track?data=ewogICJ3cml0ZUtleSI6ICJBV2NjaWd1UkhKODBuNkJ4WlI4cHRaRzBINzY0RmJObCIsCiAgInVzZXJJZCI6ICJ0b255cXVzIiwKICAiZXZlbnQiOiAiTlBPSSBIb21lcGFnZSIKfQ==
)](#)
[![Github contributors](https://img.shields.io/github/contributors/nissl-lab/npoi?logo=github&label=contributors)](https://github.com/nissl-lab/npoi/graphs/contributors)
<br />

> [!IMPORTANT]
> To ensure the long-term sustainability of this project, a monthly maintenance fee has been introduced.  This fee is required to be paid by all organizations or users of any library from this project who generate revenue. You can pay the fee via [GitHub Sponsors](https://github.com/sponsors/tonyqus).
>
> To enact this, an EULA on binary releases is added to the repo and nuget packages since 2.8.0 that requires payment of the maintenance fee.
>
> For more information on who must pay the fee and other frequently asked questions, please see the [Open Source Maintenance Fee](https://opensourcemaintenancefee.org/consumers/) organisation page.


What's NPOI
===================
This project is the .NET version of Apache POI project. With NPOI, you can read/write Office 2003/2007 files very easily.<br />

Documentation
============

[NPOI Official Tutorials](https://www.youtube.com/@codingwithtonyqus)

[Getting Started with NPOI](https://github.com/nissl-lab/npoi/wiki/Getting-Started-with-NPOI)

[How NPOI library succeeds](https://tonyqus.medium.com/how-npoi-succeeds-67ceb7333eb)

[How to use NPOI on Linux](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)

[Useful Packages based on NPOI](https://github.com/nissl-lab/npoi/wiki/ORM-on-NPOI)

Sponsors
=================
Thank you for all the sponsors. [Here is the sponsors list](Sponsors.md) 

Advantage of NPOI
=================
a. It's totally free to startups, freelancers, and hobbyists (whose annual revenue is less than $10,000 USD)

b. Cover most features of Excel (cell style, data format, formula and so on)

c. Supported formats: xls, xlsx, docx.

d. Designed to be interface-oriented (take a look at NPOI.SS namespace)

e. Support not only export but also import

f. Real successful cases all over the world

g. Works on both Windows and Linux 

PPTX foundation (XSLF)
======================
The initial `NPOI.XSLF.UserModel` API supports creating an empty transitional
`.pptx` presentation, adding blank slides, and saving and reopening the package:

```csharp
using NPOI.XSLF.UserModel;
using System.IO;

using var presentation = new XMLSlideShow();
presentation.CreateSlide();
presentation.CreateSlide();
using var output = File.Create("blank-slides.pptx");
presentation.Write(output);
```

Use `new XMLSlideShow(inputStream)` or `new XMLSlideShow(opcPackage)` to open an
existing presentation, and `GetSlides()` to access slides in presentation order.
Closing the presentation closes its underlying OPC package. As with the other
OOXML APIs, save to a different file when opening a package from a file path.

The foundation follows the existing OOXML layering: partial PresentationML data
models live in `OpenXmlFormats/Presentation`, while `XMLSlideShow`, `XSLFSlide`,
`XSLFRelation`, and `XSLFFactory` live in `ooxml/XSLF/UserModel`. Packaging uses the
existing OPC implementation. Slides, masters, and layouts share `XSLFSheet`.
`GetSlideMasters()` and `GetSlideLayouts()` follow their XML reference-list order.
`XSLFTheme` exposes read-only theme metadata using the existing DrawingML model.
Unchanged master, layout, and theme parts retain their original bytes; unmodeled
content remains intact when supported sheet models are saved.

Use a layout from the same presentation to create a slide explicitly:

```csharp
var master = presentation.GetSlideMasters()[0];
var layout = master.GetLayout("Blank");
if (layout == null)
    throw new InvalidOperationException("This master has no layout named Blank.");
var slide = presentation.CreateSlide(layout);
var theme = slide.GetTheme(); // May be null when the master has no theme.
```

Layouts expose `Name` and the PresentationML `Type` string. Slides expose
`GetSlideLayout()`, `GetSlideMaster()`, and `GetTheme()`. The layout must be listed
by a referenced master in the same presentation. Attaching a layout does not
yet copy placeholder shapes or resolve inherited text formatting. Theme editing
is not supported.

Slide lifecycle and page size
----------------------------
`SetSlideOrder(slide, newIndex)` moves a slide to a zero-based position.
`RemoveSlide(index)` removes it from the presentation and returns its wrapper.
Remaining slide IDs and relationship IDs are preserved; the extended-properties
slide count is refreshed when saving. Direct low-level slide-list edits must not
leave the user model out of sync before calling these operations.

Removal does not recursively delete layouts, masters, themes, notes, charts,
media, or other dependencies. The removed slide part is also retained if another
package relationship still references it, including a notes back-reference or
an internal link. Changes made through the removed wrapper are not saved.
This conservative behavior can leave orphan parts and does not reclaim space.

Removal is rejected if a custom show or section references the target slide.
Reordering sectioned presentations is not supported yet. Opaque presentation
extensions or alternate content can also block ordering/removal rather than
risking invalid references. Custom-show order is independent of normal slide
order and is not changed by `SetSlideOrder`.

`GetPageSize()` and `SetPageSize(SkiaSharp.SKSize)` use points (1/72 inch), not
pixels. Each dimension must be finite and between 72 and 4032 points inclusive
(914400 through 51206400 EMUs). Setting dimensions selects the `custom` size
type, without scaling shapes or changing notes dimensions:

```csharp
presentation.SetPageSize(new SkiaSharp.SKSize(960, 540));
presentation.SetSlideOrder(presentation.GetSlides()[1], 0);
presentation.RemoveSlide(1);
```

Shape containers
----------------
Slides, layouts, and masters implement `IXSLFShapeContainer`. `GetShapes()`
returns a read-only snapshot of immediate children in XML drawing order.
`XSLFSimpleShape` wraps existing `p:sp` elements, including text-bearing shapes;
`XSLFGroupShape` exposes nested children through the same container API.
Existing identifiers remain unchanged. New IDs reserve the sheet's existing
PresentationML nonvisual identifiers, including nested and alternate content.
Repeated IDs already present in unsupported alternatives are not rewritten.

```csharp
var slide = presentation.CreateSlide();
var shape = slide.CreateAutoShape(NPOI.OpenXmlFormats.Dml.ST_ShapeType.rect);
shape.Name = "Rectangle";
shape.SetAnchor(SkiaSharp.SKRect.Create(36, 36, 180, 90));

var group = slide.CreateGroup();
group.SetAnchor(SkiaSharp.SKRect.Create(300, 36, 240, 180));
group.SetInteriorAnchor(SkiaSharp.SKRect.Create(0, 0, 120, 90));
var child = group.CreateAutoShape(NPOI.OpenXmlFormats.Dml.ST_ShapeType.ellipse);
child.SetAnchor(SkiaSharp.SKRect.Create(12, 12, 60, 45));
```

`GetAnchor()` / `SetAnchor(SKRect)` and the group's `GetInteriorAnchor()` /
`SetInteriorAnchor(SKRect)` use points (1/72 inch), not pixels. A group's child
anchors are in its local coordinate system, not absolute slide coordinates.
Exterior and interior anchors are independent: neither setter changes child
coordinates, scales shapes, nor normalizes rotation/flips. New shapes/groups
start with empty anchors; set both group rectangles explicitly for scaling.
Positions may be negative, but extents must be nonnegative and every component
must be finite and fit the DrawingML coordinate range. Conversion rounds to the
nearest EMU using 64-bit values; `SKRect` itself has single-precision coordinates.
Positions must fit -27273042329600 through 27273042316900 EMUs; each extent
must fit 0 through 2147483647 EMUs. Position and extent ranges are not identical.
Only explicit numeric EMU transforms are read. Missing/inherited transforms
and universal-measure strings are not resolved and fail explicitly rather
than returning guessed coordinates. An explicit anchor can be set when absent.
`ShapeType` returns null for a simple shape without preset geometry; custom
geometry is retained, not interpreted. The existing enum's non-schema `none`
and `lineInv` values are not accepted for preset creation.

Pictures, tables, connectors, alternate content, and unfamiliar drawing
elements are exposed as `XSLFUnsupportedShape` and preserved without mutation.
Their common ID/name can be read only when directly identifiable; anchor and
name edits are not supported. Supported edits retain existing text, styles,
unknown drawing XML, and relationship parts. Live XML is available through
`GetXmlObject()`, but low-level callers remain responsible for consistency.
Read-only/disposed presentations, detached wrappers, and ambiguous supported
XML reject high-level edits. Shape removal, copying, placeholder inheritance,
text editing, fill/line editing, and rendering are not included in this stage.

This is a foundation, not full PowerPoint support. Parameterless slide creation
requires an existing blank layout (included in newly created presentations).
The layout overload can select other layout types. Existing slide
content and unsupported related parts are retained, but high-level text boxes,
pictures, tables, charts, notes, rendering, and animation APIs are not yet
implemented. Strict OOXML and macro-enabled presentation formats are not supported.

System Requirement
===================
.NET 8

.NET Standard 2.1 (.NET Core 3.x, .NET 5, .NET 6)

.NET Standard 2.0 (.NET Core 2.x)

.NET Framework 4.0 and above

Contributors
===================
![Contributors](https://contrib.rocks/image?repo=nissl-lab/npoi)

<a href="https://info.flagcounter.com/bndt"><img src="https://s01.flagcounter.com/count/bndt/bg_FFFFFF/txt_000000/border_CCCCCC/columns_3/maxflags_20/viewers_0/labels_1/pageviews_1/flags_0/percent_0/" alt="Flag Counter" border="0"></a>
