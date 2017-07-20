# Defines the expression factory for any expressions we want
# to store

try:
    import re

    from factory import DataFactory
    
    from lexer import (
        TokenMatcher, TokenPattern,
        NumericToken, StringToken, OperatorToken,
        VariableToken, KeywordToken, FunctionToken
    )
    
    # This is a global variable to hold expressions so we only have
    # to build them once and can reuse multiple times (helps improve
    # the efficiency when running large numbers of checks on multiple
    # lines of data)
    exprFactory = DataFactory(data={
        # Matches math expressions
        'math': TokenMatcher(pattern=[
            NumericToken,
            TokenPattern(OperatorToken, [('value', '(?:[^=])'),], as_regex=True),
            NumericToken
        ]),
        # Regex to pull appart simple math expressions
        're.math': re.compile(r'((?:-?[\d]+(?:[\s]+)?[-+\/*\\](?:[\s]+)?)+(?:-?[\d]+))'),
        # Matches string concatenations
        'string.concatenate': TokenMatcher(pattern=[
            StringToken,
            TokenPattern(OperatorToken, [('value', '(?:[\+\&])'),], as_regex=True),
            StringToken
        ]),
        'function.prototype': TokenMatcher(pattern=[
            (KeywordToken, [('value', '(?:Sub|Function)'),], True),
            FunctionToken
        ]),
        'function.end': TokenMatcher(pattern=[
            (KeywordToken, [('value', 'End'),]),
            (KeywordToken, [('value', '(?:Sub|Function)'),], True),
        ]),
        'variable.assignment': TokenMatcher(pattern=[
            VariableToken,
            (OperatorToken, [('value', '='),])
        ]),
        'variable.declaration': TokenMatcher(pattern=[
            TokenPattern(KeywordToken, [('value', '(?:Re)?(?:Dim)'),], as_regex=True),
            VariableToken
        ]),
    })
except ImportError as e:
    print("{0}: {1}".format(__name__, e))