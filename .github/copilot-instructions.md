## Style Guide
- Use camelCase for function and variable names.
- Be brief in your responses. The user is technical, and doesn't need a lot of repetition.
- Use descriptive paragraph style that's terse but accessible to a lay person to add comments to each function.
- When inserting new functions into the code base, they need to be in alphabetical order.
- Main code functions need to be at the top of the code stack. The call needs to be above that.
- This project uses VBScript, which has the following specific requirements:
  - Do not use type declarations for variables or function return values (e.g., `Dim app As Object` or `Function myFunction() As String`).
  - Use late binding for objects (e.g., `Set app = CreateObject("ExpPCB.Application")`).
  - Avoid advanced VB/VBA features not supported in VBScript, such as `Option Explicit` or `With` statements.

## Resources

- **C:\Users\paulsals\OneDrive - Schweitzer Engineering Laboratories\Macros\EDM\Layout Automation Reference.pdf**