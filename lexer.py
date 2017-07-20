import re
import inspect


class Token(object):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return self.value
        
    def __repr__(self):
        return self.__str__()
        
    @classmethod
    def test(cls, expr, value):
        if type(expr) == str:
            expr = re.compile(expr)
            
        return expr.match(value) is not None

class KeywordToken(Token):
    def __init__(self, value):
        super(KeywordToken, self).__init__(value)


class FunctionToken(Token):
    # Match function calls with parenthesis
    func_expr = re.compile(r'^(?P<name>[^\(]+)(?:\()(?P<args>[^\)$]*?)(?:\))$')
    
    def __init__(self, value):
        super(FunctionToken, self).__init__(value)
        
    @classmethod
    def create(cls, name, arguments):
        return cls('{0}({1})'.format(
            name,
            [' '.join([a.__str__() for a in arg]) for arg in arguments]
        ))
        
    @property
    def name(self):
        m = self.func_expr.match(self.value)
        return m.group('name')
        
    @property
    def arguments(self):
        m = self.func_expr.match(self.value)
        
        # We will attempt to split by ','. If the naive approach fails,
        # we will repair it automatically with our loop below.
        args_ary = m.group('args').split(',')
        
        lexer = VBLexer()
        
        pArg = 0
        pOffset = 1
        
        # This is an array containing all of the arguments
        # Since arguments can be complex, the array is a list of
        # token lists (each token list represents 1 argument)
        args = []
        
        while (pArg + pOffset) < (len(args_ary) + 1):            
            tokens = lexer.tokenize(','.join(args_ary[pArg:pArg + pOffset]))
            
            # Make sure we aren't just the base token...
            # We should be some complex token (Numeric/Variable/String/etc)
            if len(tokens) > 0 and Token not in [t.__class__ for t in tokens]:
                args.append(tokens)
                pArg = pArg + pOffset
                pOffset = 0
                
            pOffset += 1
        
        # Take anything we struggled to parse during the while loop and just
        # let the lexer handle the default Token creation we will use
        if pArg < len(args_ary):
            tokens = lexer.tokenize(','.join(args_ary[pArg:]))
            
            if len(tokens) > 0:
                args.append(tokens)
        
        # If we couldn't find any arguments, we just create a default Token
        # for all of the args wrapped so we don't lose the data
        if len(args) == 0:
            # Need to use a nested loop to be consistent
            # with how arg parsing is working
            args = [[Token(m.group('args'))]]
        
        return args

    # Note: We could just use self.value if we wanted and get almost
    # identical results for this
    def __str__(self):
        return '{0}({1})'.format(
            self.name, ', '.join(
                [' '.join([a.__str__() for a in arg]) for arg in self.arguments]
            )
        )
        
class NumericToken(Token):
    def __init__(self, value):
        super(NumericToken, self).__init__(value)
        
        
class StringToken(Token):
    def __init__(self, value):
        super(StringToken, self).__init__(value)

    @property
    def data(self):
        # Return the string without the outside quotes
        return self.value[1:len(self.value) - 1]
        
        
class VariableToken(Token):
    def __init__(self, value):
        super(VariableToken, self).__init__(value)
        
        
class CommentToken(Token):
    def __init__(self, value):
        super(CommentToken, self).__init__(value)
        

class OperatorToken(Token):
    def __init__(self, value):
        super(OperatorToken, self).__init__(value)
 
        
class Lexer(object):
    base_token = Token
    default_token = Token

    def __init__(self):
        self.operators = []
        self.token_types = []

    def valid_token(self, line, start, offset):
        # This is our data we are interested in
        data = line[start:start + offset]
        before = line[:start]
        after = line[start + offset:]
        
        for token,config in self.token_types:
            if token.test(config['expr'], data):                
                # Check our before case if we have one
                if 'before' in config:
                    if not config['before'].search(before):
                        continue
                # Check our after case if we have one
                if 'after' in config:
                    # If this fails, we just need to increase our offset
                    # and move on (we must be a different token type)
                    if not config['after'].search(after):
                        continue
                # Passed all checks (before/after/expr match)
                return token,config
                
    def tokenize(self, line):
        # Helper function to tokenize a list of input lines
        # instead of having the caller to be required to 
        # manually invoke tokenize in a loop
        def __tokenize_list__(__data__):
            __tokenized_data__ = []
            
            for __line__ in __data__:
                __tokenized_data__.append(self.tokenize(__line__))
            
            return __tokenized_data__
        
        if isinstance(line, list):
            return __tokenize_list__(line)
        
        # This array will be an ordered array of the tokens
        # in the order that we encountered them
        tokens = []

        # If our line isn't a string, we just return the empty tokens list
        if not isinstance(line, str):
            return tokens

        # Removing leading/trailing white space
        # White space doesn't matter for VBA
        line = line.strip()
            
        s = ""
        tstart = 0
        toffset = 1
        
        while (tstart + toffset) < len(line):
            # Skip whitespace if our token string is empty or if we are at the start of a comma
            # since comma's would only be seen in function calls as the "first" character in a token
            if line[tstart:tstart + toffset].isspace() or line[tstart:tstart + toffset] == ',':
                tstart += 1
                continue
            try:
                tok,conf = self.valid_token(line, tstart, toffset)
            
                # No valid tokens yet, we need to just increase the offset
                if tok is None:
                    toffset += 1
                    continue
                
                # Check if we need to peek further
                if line[tstart:tstart + toffset] != line[tstart:tstart + toffset + 1]:
                    # we need to peek at the next one to see if our token is continuing
                    if self.valid_token(line, tstart, toffset + 1) is None:                    
                        tokens.append(tok(line[tstart:tstart + toffset]))
                        tstart = tstart + toffset
                        toffset = 0 # It will be increased to 1 with the += call below
            except TypeError:
                # If we made it here, we couldn't unpack tok,conf
                # Just let the toffset work as planned below
                pass
                
            toffset += 1
        
        if tstart < len(line):
            try:
                # We are still building a token so let's try to get the final result!
                tok,conf = self.valid_token(line, tstart, offset=(len(line) - tstart))
            except:
                # Just use the default_token class
                tok = self.default_token
            
            tokens.append(tok(line[tstart:]))
            
        return tokens

    def untokenize(self, data):
        try:
            # Either an empty list or data was None
            if not data:
                return ''
        
            return ' '.join([t.__str__() for t in data])
        except:
            return None        
            
    def dumps(self, token_list, file=None):
        # Helper function to identify block keywords
        def __get_block_keyword__(__line__):
            # Try the single keywords first
            try:
                __token__ = __line__[0].value
                
                if __token__ in self.block_keywords:
                    return self.block_keywords[__token__]
                    
                # Let's try combination keywords (some have 2 words)
                __token__ = ' '.join([__t__.value for __t__ in __line__[0:2]])
                
                if __token__ in self.block_keywords:
                    return self.block_keywords[__token__]
                
            except IndexError:
                pass
                
            return None
        
        def __adjust__(__var__, __num__, minimum=None, maximum=None):
            __result__ = __var__ + __num__

            # Don't let the result go lower than "minimum"
            if minimum is not None:
                __result__ = max(__var__ + __num__, minimum)
                
            # Don't let the result go higher than maximum
            if maximum is not None:
                __result__ = min(__var__ + __num__, maximum)
            
            return __result__
        
        def __is_tokenized_code__(__data__):
            try:
                if not isinstance(__data__, list):
                    return False
                    
                for __line__ in __data__:
                    if len(__line__) == 0:
                        continue
                    
                    for __tok__ in __line__:
                        if not isinstance(__tok__, self.base_token):
                            return False
            except:
                return False
            
            # Nothing failed the testing, we must be tokenized code
            return True
        
        if not __is_tokenized_code__(token_list):
            raise TypeError('Error - token_list is not valid tokenized code')
        
        data = []
        tabs = 0
        
        for line in token_list:
            if len(line) == 0:
                continue
            
            block = __get_block_keyword__(line)
            
            if block is not None:
                tabs = __adjust__(tabs, block['before'], minimum=0)
                
            data.append("{0}{1}".format('\t' * tabs, self.untokenize(line)))
            
            if block is not None:
                tabs = __adjust__(tabs, block['after'], minimum=0)
                
        if file is not None:
            with open(file, "w") as f:
                f.write('\n'.join(data))
        
        return data
        
    
class VBLexer(Lexer):
    def __init__(self):
        self.operators = [
            '=', '&', '&=', '*', '*=', '/', '/=', '\\', '\=', '^',
            '^=', '+', '+=', '-', '-=', '>>', '>>=', '<<','<<=', '<>'
        ]
        
        self.keywords = [
            'AddHandler', 'AddressOf', 'Alias', 'And', 'AndAlso', 'As', 'Boolean',
            'ByRef', 'Byte', 'ByVal', 'Call', 'Case', 'Catch', 'CBool', 'CByte', 'CChar',
            'CDate', 'CDec', 'CDbl', 'Char', 'CInt', 'Class', 'CLng', 'CObj', 'Const',
            'Continue', 'CSByte', 'CShort', 'CSng', 'CStr', 'CType', 'CUInt', 'CULng',
            'CUShort', 'Date', 'Decimal', 'Declare', 'Default', 'Delegate', 'Dim', 'DirectCast',
            'Do', 'Double', 'Each', 'Else', 'ElseIf', 'End', 'EndIf', 'Enum', 'Erase', 'Error',
            'Event', 'Exit', 'False', 'Finally', 'For', 'Friend', 'Function', 'Get', 'GetType',
            'GetXMLNamespace', 'Global', 'GoSub', 'GoTo', 'Handles', 'If', 'Implements',
            'Imports', 'Imports', 'In', 'Inherits', 'Integer', 'Interface', 'Is', 'IsNot',
            'Let', 'Lib', 'Like', 'Long', 'Loop', 'Me', 'Mod', 'Module', 'MustInherit',
            'MustOverride', 'MyBase', 'MyClass', 'Namespace', 'Narrowing', 'New', 'Next', 'Not',
            'Nothing', 'NotInheritable', 'NotOverridable', 'Object', 'Of', 'On', 'Operator',
            'Option', 'Optional', 'Or', 'OrElse', 'Overloads', 'Overridable', 'Overrides',
            'ParamArray', 'Partial', 'Private', 'Property', 'Protected', 'Public', 'RaiseEvent',
            'ReadOnly', 'ReDim', 'REM', 'RemoveHandler', 'Resume', 'Return', 'SByte', 'Select',
            'Set', 'Shadows', 'Shared', 'Short', 'Single', 'Static', 'Step', 'Stop', 'String',
            'Structure', 'Sub', 'SyncLock', 'Then', 'Throw', 'To', 'True', 'Try', 'TryCast',
            'TypeOf', 'Variant', 'Wend', 'UInteger', 'ULong', 'UShort', 'Using', 'When', 'While',
            'Widening', 'With', 'WithEvents', 'WriteOnly', 'Xor', '#Const', '#Else', '#ElseIf',
            '#End', '#If'
        ]
        
        self.block_keywords = {
            'Function': { 'before': 0, 'after': 1 },
            'Sub':      { 'before': 0, 'after': 1 },
            'End':      { 'before': 0, 'after': 1 },
            'If':       { 'before': 0, 'after': 1 },
            '#If':      { 'before': 0, 'after': 1 },
            'Else':     { 'before': -1, 'after': 1 },
            '#Else':    { 'before': -1, 'after': 1 },
            'ElseIf':   { 'before': -1, 'after': 1 },
            'End' :     { 'before': -1, 'after': 0 },
            '#End' :    { 'before': -1, 'after': 0 },
            'Select':   { 'before': 0, 'after': 1 },
            'Case':     { 'before': -1, 'after': 1},
            'For':      { 'before': 0, 'after': 1 },
            'Next':     { 'before': -1, 'after': 0 },
            'Do':       { 'before': 0, 'after': 1},
            'While':    { 'before': 0, 'after': 1},
            'Loop':     { 'before': -1, 'after': 0},
            'Wend':     { 'before': -1, 'after': 0},
            'Private Sub':
                        { 'before': 0, 'after': 1 },
            'Public Sub':
                        { 'before': 0, 'after': 1 },
            'Private Function':
                        { 'before': 0, 'after': 1 },
            'Public Function':
                        { 'before': 0, 'after': 1 },
        }
        
        self.token_types = [
            (KeywordToken, {
                'before': re.compile(r'(?:^|[\s]+)$'),
                'expr': re.compile(r'^(?:{0})$'.format('|'.join(self.keywords)), re.IGNORECASE),
                'after': re.compile(r'^([\s]+|$)'),
            }),
            (VariableToken, {
                # This variable token is to match for array initializations:
                # e.g. Dim loathe(53) As Long
                # Check array's before we check functions so the lexer doesn't get confused
                'before': re.compile(r'(?:(?:Re)?Dim[\s]+)(?:Preserve[\s]+)?$', re.IGNORECASE),
                # varname(number)
                'expr': re.compile(r'^(?:[a-z])(?:[^\.\!\@\&\$\#\s\(\,]{0,254})(?:\((?:[\d]+)(?:,[\d]+)*?\))$', re.IGNORECASE),
            }),
            (FunctionToken, {
                # TODO: Fix this function regex to handle functions that don't use ()'s
                # Might be able to just add a seperate FunctionToken regex to this list
                # Might need to add a (?=) lookahead for VariableToken to make sure there isn't
                # a list of arguments like a non-()'s function call
                'expr': re.compile(r'^(?P<name>[^\(]+)\((?P<args>[^\)]*?)\)$'),
            }),
            (NumericToken, {
                'expr': re.compile(r'^(?:[\d]+)(?:(?=\.)\.[\d]+)?$'), # Negative numbers = just another subtraction
                'after': re.compile(r'^(?:[^\.]|$)'), # Cannot have a period immediately after (that would be a float)
            }),
            (StringToken, {
                'expr': re.compile(r'^(["])(?:(?=(\\?))\2.)*?\1$'),
            }),
            (VariableToken, {
                # Using a negative lookahead at the end of the expr to make sure that there
                # is not an open paren indicating a function call
                'expr': re.compile(r'^(?:[a-z])(?:[^\.\!\@\&\$\#\s\(\,]{0,254})(?!\()$', re.IGNORECASE),
                'after': re.compile(r'^(?:[{0}\s,]|$)'.format(''.join(['\\' + op for op in self.operators]))),
            }),
            (CommentToken, {
                'expr': re.compile(r'^(?:\'.*?)$'),
            }),
            (OperatorToken, {
                'expr': re.compile(r'^(?:{0})$'.format('|'.join(['\\' + op for op in self.operators]))),
            })
        ]

        
class TokenMatch(object):
    def __init__(self, data, span):
        self.data = data
        self.span = span

    @property
    def match(self):
        start,end = self.span
        return self.data[start:end]
        
    def __repr__(self):
        return self.match
        
    
   
class TokenMatcher(object):
    def __init__(self, pattern):
        self.pattern = []
        
        for item in pattern:
            try:
                self.pattern.append(TokenPattern.compile(item))
            except TypeError as e:
                raise TypeError('TokenMatcher: {0}'.format(e))
                
    # Helper function used by match/matches
    def __test_sample__(self, __sample__):
        # Look through our pattern and test them against the sample
        for pTok, sTok in zip(self.pattern, __sample__):
            if not pTok.test(sTok) and not pTok.optional:
                return False
                
        return True

    # This will test the input data start from the beginning
    # Use 'search' to get matches that can be offset from the 
    # beginning of the input data        
    def match(self, data):
        poffset = len(self.pattern)
        
        if poffset > len(data):
            return None
        
        # Test our sample to see if we get a match
        if self.__test_sample__(data[:poffset]):
            return TokenMatch(data, (0, poffset))
            
        # No matches found
        return None
        
    # Use to return only the first match found
    def search(self, data):
        # Get the length of our pattern
        poffset = len(self.pattern)
        
        # Can't match if our pattern is longer than our data
        if poffset > len(data):
            return None
        
        # Try to go through all of our data and get the
        # "subtoken" arrays to match our pattern
        for pstart in range(0, len(data) - poffset + 1):
            sample = data[pstart:pstart + poffset]
            # Test the sample and return the first match
            if self.__test_sample__(sample):
                return TokenMatch(data, (pstart, pstart + poffset))
        
        return None
         
    # Use to return all matches found
    def searches(self, data):                    
        # Get the length of our pattern
        poffset = len(self.pattern)
        
        # Can't match if our pattern is longer than our data
        if poffset > len(data):
            return None
            
        matches = []
        
        # Try to go through all of our data and get the
        # "subtoken" arrays to match our pattern
        for pstart in range(0, len(data) - poffset + 1):
            sample = data[pstart:pstart + poffset]
            # Test the sample and return the first match
            if self.__test_sample__(sample):
                matches.append(TokenMatch(data, (pstart, pstart + poffset)))
        
        if len(matches) == 0:
            return None # No matches... :(

        return matches
        
        
class TokenPattern(object):
    def __init__(self, cls, key_value_pairs=None, as_regex=False, optional=False):
        # class that the object should be
        self.cls = cls

        # this should be a list of tuples [(key,value), ..]
        self.key_value_pairs = key_value_pairs
        
        # Specifies if the kvp's should be compared using regex or not
        self.as_regex = as_regex
        
        # Let's the match know if this pattern is optional or not
        self.optional = optional
        
    def __str__(self):
        return '{0}({1})'.format(self.cls.__name__, self.key_value_pairs)
        
    def test(self, data):        
        # Not an instance of our class so it fails
        if not isinstance(data, self.cls):
            return False
            
        if self.key_value_pairs:
            for key, value in self.key_value_pairs:
                # Doesn't have the key so it fails
                if not hasattr(data, key):
                    return False
                
                # Get the value for our key
                v = getattr(data, key)
                
                # Check value normally if not using regex, else
                # we check it with regex using search
                if (not self.as_regex and (v != value)) or \
                        (self.as_regex and (not re.search(value, v))):
                    return False
        
        # None of our tests failed so we are good to go!
        return True        

    @classmethod
    def compile(cls, pattern):
        if isinstance(pattern, cls):
            return pattern
        elif isinstance(pattern, tuple):
            return cls(*pattern)
        elif inspect.isclass(pattern):
            return cls(cls=pattern)
        else:
            raise TypeError("Unknown pattern: {0}".format(pattern))
        
        
def print_tokens(tokens):
    for token in tokens:
        if isinstance(token, FunctionToken):
            print('> {0} :: {1}'.format(type(token).__name__, token.name))
            
            for i in range(0, len(token.arguments)):
                arg = token.arguments[i]
                print(' | var {0} ::'.format(i + 1))
                for tok in arg:
                    print('  - {0} :: {1}'.format(type(tok).__name__, tok))
        else:
            print('> {0} :: {1}'.format(type(token).__name__, token))


def main():
    lexer = VBLexer()
    
    tests = [
        "Sub streptococcal(acaritre, arrowsmith, gauche)",
        "Dim loathe(63) As Long"
    ]
    
    for test in tests:
        print("------------- TEST -------------")
        print("Input: {0}".format(test))
        print("")
        
        tokens = lexer.tokenize(test)
        print_tokens(tokens)
        print("")
    
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print('User sent keyboard interrupt - exiting')
        