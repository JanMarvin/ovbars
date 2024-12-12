test_that("loading files work", {

  fl <- "https://github.com/JanMarvin/openxlsx-data/raw/refs/heads/main/gh_issue_416.xlsm"
  wb <- openxlsx2::wb_load(fl)
  vba <- wb$vbaProject

  exp <- list(Sheet1 = "Attribute VB_Name = \"Sheet1\"\nAttribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nAttribute VB_TemplateDerived = False\nAttribute VB_Customizable = True\nPrivate Sub Worksheet_SelectionChange(ByVal Target As Range)\n    #donothing\nEnd Sub\n",
              Sheet2 = "Attribute VB_Name = \"Sheet2\"\nAttribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nAttribute VB_TemplateDerived = False\nAttribute VB_Customizable = True\n",
              ThisWorkbook = "Attribute VB_Name = \"ThisWorkbook\"\nAttribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nAttribute VB_TemplateDerived = False\nAttribute VB_Customizable = True\n")
  got <- ovbars::ovbar_out(name = vba)
  expect_equal(exp, got)

  if (Sys.info()[["sysname"]] == "Windows") {
    exp <- list(PROJECT = "/PROJECT", PROJECTwm = "/PROJECTwm", `Root Entry` = "/",
                Sheet1 = "/VBA\\Sheet1", Sheet2 = "/VBA\\Sheet2", ThisWorkbook = "/VBA\\ThisWorkbook",
                VBA = "/VBA", `_VBA_PROJECT` = "/VBA\\_VBA_PROJECT", `__SRP_0` = "/VBA\\__SRP_0",
                `__SRP_1` = "/VBA\\__SRP_1", `__SRP_2` = "/VBA\\__SRP_2", `__SRP_3` = "/VBA\\__SRP_3",
                dir = "/VBA\\dir")
  } else {
    exp <- list(PROJECT = "/PROJECT", PROJECTwm = "/PROJECTwm", `Root Entry` = "/",
                Sheet1 = "/VBA/Sheet1", Sheet2 = "/VBA/Sheet2", ThisWorkbook = "/VBA/ThisWorkbook",
                VBA = "/VBA", `_VBA_PROJECT` = "/VBA/_VBA_PROJECT", `__SRP_0` = "/VBA/__SRP_0",
                `__SRP_1` = "/VBA/__SRP_1", `__SRP_2` = "/VBA/__SRP_2", `__SRP_3` = "/VBA/__SRP_3",
                dir = "/VBA/dir")

  }
  got <- ovbars::ovbar_meta(name = vba)
  expect_equal(exp, got)

})
