# VBA Grimoire: Maximize and Minimize Buttons for UserForms

Welcome to the VBA Grimoire, where we unlock the power of custom VBA scripts to enhance your Office applications. This repository contains a powerful VBA module designed to detect your server environment (32-bit or 64-bit) and appropriately add Maximize and Minimize buttons to your UserForms.

## Overview

The `MinimiseMaximiseIcons.bas` module ensures that your UserForms are equipped with Maximize and Minimize buttons, regardless of whether you're running a 32-bit or 64-bit version of Excel. By automatically detecting the environment, the module applies the correct API declarations, enabling you to maximize the flexibility and usability of your forms.

## Installation

1. Download the Module:
   - Download the `MinimiseMaximiseIcons.bas` file from this repository.

2. Import into Your VBA Project:
   - Open your Office application (e.g., Excel) and press `ALT + F11` to open the VBA editor.
   - In the Project Explorer, right-click on your VBA project and choose `Import File...`.
   - Select the `MinimiseMaximiseIcons.bas` file you downloaded and click `Open`. The module will be added to your project.

3. Use the Subroutines: 
   - In your UserForm's `Initialize` event, add calls to the subroutines `AddMaximizeButton` and `AddMinimizeButton` to apply the respective buttons.

Example:
```vba
Private Sub UserForm_Initialize()
    AddMaximizeButton Me
    AddMinimizeButton Me
End Sub
```

4. Compile with `Option Explicit`: Ensure that `Option Explicit` is declared at the top of your modules to prevent errors and ensure all variables are properly defined.

## Usage

- AddMaximizeButton(UserForm As Object): Adds a Maximize button to the specified UserForm.
- AddMinimizeButton(UserForm As Object): Adds a Minimize button to the specified UserForm.

## Contributing

Feel free to contribute to this repository by submitting issues or pull requests. Contributions that add new features or improve existing code are always welcome!

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.

## Contact

For any questions, feel free to reach out via the Issues section of this repository.
