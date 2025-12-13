# VBA UserForms Class Icons

English | [Русский](README_RUS.md)

A comprehensive collection of icons for VBA UserForms applications using the Segoe MDL2 Assets font. This library provides easy access to a wide range of Microsoft-style icons that can be used in Excel VBA applications.

## Features

- Over 1,30 Microsoft-style icons available
- Based on the Segoe MDL2 Assets font
- Easy-to-use VBA enum for accessing icons
- Functions to retrieve icon characters, codes, and names

## Installation

1. Import the `modIcon.bas` module into your VBA project
2. Ensure your system has the Segoe MDL2 Assets font installed (standard on Windows 10/11)

## Usage

### Basic Usage

```vba
' Get an icon character
Dim iconChar As String
iconChar = getIcon(icAccept) ' Returns the accept/checkmark icon

' Get an icon code
Dim iconCode As Long
iconCode = getIconCode(icHeart) ' Returns the Unicode code for the heart icon

' Get an icon name
Dim iconName As String
iconName = getIconName(icStar) ' Returns the string name of the icon
```

### Using Icons in UserForms

```vba
' Set icon in a label
Label1.Caption = getIcon(icAccept)
Label1.Font.Name = "Segoe MDL2 Assets"
Label1.Font.Size = 16

' Set icon in a button
Button1.Caption = getIcon(icSave)
Button1.Font.Name = "Segoe MDL2 Assets"
```

## Available Icons

The library includes the following categories of icons:

- Action icons (accept, cancel, save, delete, etc.)
- Navigation icons (back, forward, home, etc.)
- Communication icons (mail, phone, chat, etc.)
- Media icons (play, pause, video, etc.)
- Document icons (folder, file, print, etc.)
- Settings and system icons
- And many more...

See the full list in the `IconEnums` enum in the `modIcon.bas` file.

## Functions

- `getIcon(codIcon)` - Returns the icon character as a string
- `getIconCode(codIcon)` - Returns the Unicode code of the icon
- `getIconName(codIcon)` - Returns the name of the icon
- `getAllIconsArray()` - Returns an array with all available icons

## Requirements

- Microsoft Excel with VBA support
- Segoe MDL2 Assets font (included with Windows 10/11)

## License

Licensed under the Apache License, Version 2.0. See the LICENSE file for details.