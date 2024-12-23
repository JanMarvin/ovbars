
<!-- README.md is generated from README.Rmd. Please edit that file -->

# ovbars

<!-- badges: start -->

[![R-CMD-check](https://github.com/JanMarvin/ovbars/workflows/R-CMD-check/badge.svg)](https://github.com/JanMarvin/ovbars/actions)
[![r-universe](https://janmarvin.r-universe.dev/badges/ovbars)](https://janmarvin.r-universe.dev/ovbars)
<!-- badges: end -->

In the Office Open XML standard, files with the `.xlsm` extension
indicate that the workbook contains macros. When loading these files
using R packages, the macro code is not directly visible as it is
embedded in a file named `vbaProject.bin`.

This R package serves as a wrapper around the `ovba` Rust library,
enabling the parsing of the `vbaProject.bin` file. You can extract the
`vbaProject.bin` file by unzipping the XLSM file or by using the
`openxlsx2` package, as demonstrated in the example below.

``` r
## an example xlsm file
fl <- "https://janmarvin.github.io/openxlsx-data/gh_issue_416.xlsm"

## get the path to the vba file
vba <- openxlsx2::wb_load(fl)$vbaProject

## extract the macro code
code <- ovbars::ovbar_out(name = vba)

## access the code from "Sheet1"
message(code["Sheet1"])
#> Attribute VB_Name = "Sheet1"
#> Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
#> Attribute VB_GlobalNameSpace = False
#> Attribute VB_Creatable = False
#> Attribute VB_PredeclaredId = True
#> Attribute VB_Exposed = True
#> Attribute VB_TemplateDerived = False
#> Attribute VB_Customizable = True
#> Private Sub Worksheet_SelectionChange(ByVal Target As Range)
#>     #donothing
#> End Sub
```

This view differs from what you might be accustomed to in spreadsheet
software, as it includes additional information typically used
internally by the software. Lines starting with `Attribute` are entirely
hidden.

Similarly, this project allows the extraction of additional information
regarding the file system of the OVBA file.

``` r
## additional meta information regarding the OVBA file system
ovbars::ovbar_meta(name = vba)
#> $PROJECT
#> [1] "/PROJECT"
#> 
#> $PROJECTwm
#> [1] "/PROJECTwm"
#> 
#> $`Root Entry`
#> [1] "/"
#> 
#> $Sheet1
#> [1] "/VBA/Sheet1"
#> 
#> $Sheet2
#> [1] "/VBA/Sheet2"
#> 
#> $ThisWorkbook
#> [1] "/VBA/ThisWorkbook"
#> 
#> $VBA
#> [1] "/VBA"
#> 
#> $`_VBA_PROJECT`
#> [1] "/VBA/_VBA_PROJECT"
#> 
#> $`__SRP_0`
#> [1] "/VBA/__SRP_0"
#> 
#> $`__SRP_1`
#> [1] "/VBA/__SRP_1"
#> 
#> $`__SRP_2`
#> [1] "/VBA/__SRP_2"
#> 
#> $`__SRP_3`
#> [1] "/VBA/__SRP_3"
#> 
#> $dir
#> [1] "/VBA/dir"
```

## Development

``` r
# document code
savvy::savvy_update()
devtools::document()
# update vendored packages
rextendr::vendor_pkgs()
```

## License

This package is licensed under the MIT license. At the moment it bundles
a patched version of the [`ovba`](https://github.com/tim-weis/ovba)
library, which is released under the MIT License Copyright (c) 2020 Tim
Weis.
