# PowerPoint VBA Bug Report

> Bug report for Ribbon event `getPressed`

## Getting started

1. Install [.NET SDK](https://dotnet.microsoft.com/en-us/download)
2. Install [vbamc](https://www.nuget.org/packages/vbamc) tool
3. Compile addin with `make`

```shell
brew install make
brew install --cask dotnet-sdk
dotnet tool install --global vbamc

cd src
make
```

### Installing the addin

Use Microsoft PowerPoint for Mac, choose **Tools** > **PowerPoint Add-ins...** dialog
to add the compiled `GetPressedAddin.ppam` file to the application.

The addin is compiled to the `src/bin` directory.


## Bug report

The `toggleButton` in Ribbon has the `getPressed` callback method which is used to determine
if a button is pressed or not.

When multiple presentations are opened in PowerPoint and we invalidate the Ribbon using the
`IRibbonUI.Invalidate()` or `IRibbonUI.InvalidateControl()` the `getPressed` callback
will be called for each of the presentation opened.

The issue is the `IRibbonControl.Context` object will always report the same `DocumentWindow`
object (the first active one) for each callback.

Instead the correct behavior should be the `Context` to reference the window in which the
actual button exists.

Without correct context we are unable to correctly track and report the pressed state
of our buttons when multiple presentations are opened.
