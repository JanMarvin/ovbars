
<!-- README.md is generated from README.Rmd. Please edit that file -->

# ovbars

<!-- badges: start -->

[![R-CMD-check](https://github.com/JanMarvin/ovbars/workflows/R-CMD-check/badge.svg)](https://github.com/JanMarvin/ovbars/actions)
[![r-universe](https://janmarvin.r-universe.dev/badges/ovbars)](https://janmarvin.r-universe.dev/ovbars)
<!-- badges: end -->

This R package allows using the `ovba` rust library to read the
`vbaProject.bin` file which is included in the `xlsm` files as defined
by the Office Open XML standard. The macro code is embedded in a binary
OVBA file. This is readable via `ovbars`.

``` r
fl <- "https://github.com/JanMarvin/openxlsx-data/raw/refs/heads/main/gh_issue_416.xlsm"
wb <- openxlsx2::wb_load(fl)
vba <- wb$vbaProject

code <- ovbars::ovbar_out(name = vba)
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

## License

This package is licensed under the MIT license. At the moment it bundles
a patched version of the [`ovba`](https://github.com/tim-weis/ovba)
library, which is released under the MIT License Copyright (c) 2020 Tim
Weis.
