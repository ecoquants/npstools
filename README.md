# npstools
National Park Service tools for R

## Install R package

```r
devtools::install_github("ecoquants/npstools")
```

## Use R package

Will include basic usage of functions here later. For now, see:

- [Reference](./reference/) help documentation on functions
- https://github.com/ecoquants/nps-veg

## Background

Below are some notes about the creation and maintenance of this R package.

### Creation of R package

1. Create an new GitHub repo on https://github.com/new
1. Create a new Git R project (File --> New Project --> Version Control --> Git --> [paste in Repository URL] & [Create Project])
1. Use `devtools::create(path=".")` to initiate a package. You can say "no" when it asks you to overwrite the RProj file.

For more, see https://github.com/isteves/r-pkg-intro.

### Update website

To update the website for the R package, update documentation and regenerate the website outputs into the `docs/` folder:

```R
devtools::document()
pkgdown::build_site()
```

#### Errors with `pkgdown::build_site()`

You may get error like this...

```
Reading 'man/find_gaps.Rd'
Error in rep(TRUE, length(x) - 1) : invalid 'times' argument
```

To fix this, be sure that all arguments in your functions are given a definition, ie next to `#' @param some-argument-name`. Then in RStudio, place your cursor inside the offending function (eg `find_gaps()` based on error message example), Code > Insert Roxygen Skeleton. This assures all arguments are listed in the documentation of the function. Then rerun:

```R
devtools::document()
pkgdown::build_site()
```

### Package dependencies

See: Dependencies in [Package basics · R packages](http://r-pkgs.had.co.nz/description.html#dependencies)
http://r-pkgs.had.co.nz/description.html#dependencies

```R
devtools::use_package("tidyverse"  , type="Imports")
devtools::use_package("lubridate"  , type="Imports")
devtools::use_package("fs"         , type="Imports")
devtools::use_package("here"       , type="Imports")
devtools::use_package("glue"       , type="Imports")
devtools::use_package("yaml"       , type="Imports")
devtools::use_package("openxlsx"   , type="Imports")
devtools::use_package("pivottabler", type="Imports")
devtools::use_package("leaflet"    , type="Imports")
```

### Vignettes

```R
devtools::use_vignette("quick-start")
devtools::document()
pkgdown::build_site()
```

### Data

```r
# setup data-raw/ folder
devtools::use_data_raw()

# populate from local
devtools::use_data(applicants, overwrite = TRUE)
```

### Conversion from VB to SQL to R

Copied Visual Basic (VB) modules in Access database to install folder `inst/accessdb_source/*.vb` for translating from strings to pseudo-SQL into `inst/accessdb_source/*.sql` and indent for readability to then translate into dplyr functions for R (see cheat sheet [Data Transformation with dplyr](https://github.com/rstudio/cheatsheets/raw/master/data-transformation.pdf)).

Used three panes:

1. VB: [Atom](https://atom.io/) editor with [language-vb](https://atom.io/packages/language-vb) package installed.
1. SQL: RStudio
1. R nps-veg: RStudio
1. R npstools: RStudio

