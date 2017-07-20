# VBSimplifier
With the surge in maldoc based phishing, obfuscated VBscript has become a real annoyance to deal with. This python script is used to automatically perform some of the common deobfuscation techniques to allow an analysts the ability to quickly retrieve useful artifcats from the sample (such as callback urls, drop binaries, etc).

Note: The tool is in the early stages and is still a bit rough around the edges, but it should spit out much more managable VB script to review.

## Usage

```
Usage: VBSimplifier.py [options]

Options:
  -h, --help            Show help message

  Deobfuscation Options:
    Toggle deobfuscation tactics below

    --all               Enable all deobfuscation techniques
    --concatenate       Concatenates strings as needed
    --math              Solves math equations
    --comments          Strips single line comments
    --str-functions     Resolves common string functions (Left, Right,
                        StrReverse, etc)

  Input Options:
    Provide the following data as input to the program

    -i INPUT, --input=INPUT
                        Input VB Script file to deobfuscate. Required

  Output Options:
    Specify the output for the program

    -o OUTPUT, --output=OUTPUT
                        Output file to write deobfuscated code
    -s, --strings       Print strings within code
    -f, --functions     Print functions within code


--------------------------------------------------------------------------

Sample Execution:
> VBSimplifier.py --input macros.txt --output result.txt --all --strings --functions
```

Tested with: python 3.4.0
