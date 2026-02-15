# Notes

## Software Requirements Specification

### Core Architecture

I'd suggest having a two module system.
The encryption library which is hardened and handles all encryption, decryption, and access level permissions.
And the [Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview) which handles exclusively frontend logic.

### Language Selection

[Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview) must be written in JavaScript or TypeScript.
The Add-In runs inside a sandboxed browser-like environment which cannot call out to local `.dll` files.
As it cannot spin up an instance of the JVM, Java is almost entirely out of the question.
The encryption library may be written in C++ which can target [WASM](https://webassembly.org/) via [Emscripten](https://emscripten.org/).
This would allow it to run directly inside Office's sandbox.

### Assumptions and Dependencies

The frontend depends on the version of the [Excel Add-in API](https://learn.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview).
That is, it must be no less than the last breaking change to a functionality critial to our frontend.
And no more than the next breaking change to a functionality critical to our frontend.
How exactly to quantify this I do not know.
