% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/000-wrappers.R
\name{ovba}
\alias{ovba}
\alias{ovbar_out}
\alias{ovbar_meta}
\title{\code{ovbars} allows to parse the OVBA file \code{vbaProject.bin} which contains
macro variables in Office Open XML files}
\usage{
ovbar_out(name)

ovbar_meta(name)
}
\arguments{
\item{name}{the path to the input file}
}
\value{
A named list with character strings. The list is sorted alphabetically.
}
\description{
\code{ovbar_out()}: Returns a list of macro sources found in the vba binary.
\code{ovbar_meta()}: Returns a list of additional meta data of directories found
in the vba binary. The latter is probably not useful for the user.
}
