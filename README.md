This repository demonstrates a subtle bug in VBScript's `IsEmpty` function. When dealing with numerical operations,  `IsEmpty` does not handle empty strings ("") as expected for numerical calculations.  The provided solution showcases a way to properly handle these cases.

## Bug Description:
The `IsEmpty` function in VBScript is designed to check if a variable is uninitialized or empty. However, when used in a mathematical context involving empty strings, it behaves unexpectedly and might lead to errors or incorrect outcomes.

## Solution:
To fix this, we modify the code to explicitly check if a variable is an empty string ("") and handle it as a numerical 0. This is more precise in the context of numerical operations and ensures the correctness of the calculation.