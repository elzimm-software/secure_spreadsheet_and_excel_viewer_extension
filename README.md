# CEN 4033 - Team 1 - Term Project - Secure Spreadsheet and Excel Viewer Extension

## Project Description

The goal of this project is to design and develop: (a) software to generate a Secure Data
Container (SDC) as an encrypted spreadsheet file; (b) SDC viewer in the form of Microsoft Excel Add-
on.
Each SDC worksheet needs to be encrypted with a separate AES symmetric key. SDC must store access
control matrix that defines which role can access what worksheet. Authorized users must be able to view
decrypted worksheets the user is authorized for in Microsoft Excel. The encrypted spreadsheet file shall
contain at least five worksheets. Access control policies must be written in the form of Access Control
Matrix and support four roles: administrator, privileged user, user, guest. Administrator shall be able to
access all the worksheets in the spreadsheet file, privileged user shall be able to access at least three
worksheets in the spreadsheet file, user shall be able to access two worksheets in the spreadsheet file,
guest shall be able to access only one worksheet in the spreadsheet file.
Viewer shall be able to authenticate the user, determine their role (four roles must be supported), determine
which worksheets are accessible for this role, according to the access control matrix in SDC, decrypt and
display in Microsoft Excel the worksheets the user is authorized for. The inaccessible worksheets must be
hidden for users based on their role and access control policies. The viewer must prevent unauthorized
users from changing the access control matrix such that they cannot escalate their privileges. Users shall
not be able to modify the viewer’s source code, such that users cannot gain unauthorized access.

## Useful Links

- [Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview)
- [Excel Add-in API](https://learn.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview)
- [WASM](https://webassembly.org/)
- [Emscripten](https://emscripten.org/)
- [TypeScript](https://www.typescriptlang.org/)
- [C++ Static Library](https://learn.microsoft.com/en-us/cpp/build/walkthrough-creating-and-using-a-static-library-cpp?view=msvc-170)

## Relevant Works

[D. Ulybyshev et al., "Protecting Electronic Health Records in Transit and at Rest," 2020 IEEE 33rd
International Symposium on Computer-Based Medical Systems (CBMS), Rochester, MN, USA, 2020, pp.
449-452, doi: 10.1109/CBMS49503.2020.00091.](https://ieeexplore.ieee.org/abstract/document/9183319)
