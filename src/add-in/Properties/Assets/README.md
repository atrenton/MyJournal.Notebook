# OneNote Page Title Background Image

<picture>
  <img alt="SVG image" src="retro-title-background.svg" width="50%" />
</picture>

## Objectives

- Image replicates **OneNote 2010** page title background

- Image dimensions are sized to 2x scale to support 200% zoom of raster image

## Programmer Notes

- Unit of measurement in **OneNote** XML schema is Points (pt), where 0.75pt == 1 pixel (px) at 96 DPI

- **OneNote** page title viewport default size is approximately 223pt wide x 48pt high

- SVG rect `ry` attribute value is 50% of `height` attribute value

- SVG rect `rx` attribute value is 50% of `ry` value

- SVG image is dynamically sized (based on title string width) and converted to PNG raster-based format using [Svg.NET](https://github.com/svg-net/SVG "GitHub - svg-net/SVG")

- NuGet Package Dependency: [Svg 3.4.4](https://www.nuget.org/packages/Svg/3.4.4 "NuGet Gallery | Svg 3.4.4")

<!-- GitHub Docs: Basic writing and formatting syntax -->
<!-- https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax -->
