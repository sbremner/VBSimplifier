import re
import routine
import factory

from utils import (
    md5sum, lazy, combine,
)

from lexer import (
    VBLexer,
    VariableToken, FunctionToken, StringToken, Token,
)

from expressions import (
    exprFactory
)
    
from VBStack import (
    VBStack
)
    
# ---------------------------------------------------------------------------------- #

# This factory holds lazy resolutions to our routine's that are implemented below
routineFactory = factory.DataFactory(data={
    # Deobfuscation tactics:
    'comments': lazy('strip_comments_routine'),
    'math': lazy('math_routine'),
    'str_functions': lazy('string_function_routine'),
    'concatenate': lazy('merge_strings_routine'),
    'resolve': lazy('resolve_variables_routine'),
    
    # Output information:
    'strings': lazy('print_strings_routine'),
    'functions': lazy('print_functions_routine'),
})

# ---------------------------------------------------------------------------------- #

def strip_comments_matcher(line):
    line = line.strip()
    return line.startswith("'")

    
strip_comments_routine = routine.Routine(
    matcher=strip_comments_matcher,
    skip_on_match=True,
    type=routine.PRE_ROUTINE
)

# ---------------------------------------------------------------------------------- #

# This is used to simplified mathematical expressions
# Example:
#   In :: garbage = 5 + 2 - 7
#  Out :: garbage = 0
# EXPR:
#   Group 1 = variable name
#   Group 2 = math expression
def math_matcher(line):
    lexer = VBLexer()
    tokens = lexer.tokenize(line)
    
    mathexpr = exprFactory.get('math')
    
    # Return true if we have a match!        
    return mathexpr.search(tokens) != None
    
    
def math_handler(line):
    # Use regex for this math equation to preserver order of operations
    # NOTE: This regex will fail order of operations if we have special
    # operators such as >> or <<
    mathexpr = exprFactory.get('re.math')
    m = mathexpr.search(line)
    
    # Replace single backslash with double forward slash (integer division)
    equation = m.group(1).replace('\\', '//')
    
    return line.replace(m.group(0), '{0}'.format(eval(equation)))

math_routine = routine.Routine(
    matcher=math_matcher,
    handler=math_handler,
    type=routine.ROUTINE|routine.POST_ROUTINE
)

# ---------------------------------------------------------------------------------- #

# VBS string functions (Left, Right, UCase, Mid)
# Note: Regex used is very basic and can fail if ')' is used in the string
string_function_expr = re.compile(r'(Left|Right|LCase|UCase|Mid|StrReverse)\(("[^"]+"[^\)]*)\)', re.IGNORECASE)

def string_function_matcher(line):   
    return string_function_expr.search(line) != None
    
def string_function_handler(line):    
    def __left__(data, num):
        num = int(num)        
        return data[:num]
        
    def __right__(data, num):
        num = int(num)
        return data[-num:]
        
    def __lcase__(data):
        return data.lower()
        
    def __ucase__(data):
        return data.upper()
        
    def __strreverse__(data):
        return data[::-1]
        
    def __mid__(data, start, num):
        start = int(start)
        num = int(num)
        
        # Need to re-index our start since VBS is 1-based
        start = start - 1
        
        return data[start:start+num]
        
    actions = {
        "left": __left__,
        "right": __right__,
        "ucase": __ucase__,
        "lcase": __lcase__,
        "strreverse": __strreverse__,
        "mid": __mid__,
    }
    
    m = string_function_expr.search(line.strip())
    try:
        result = actions[m.group(1).lower()](*([s.strip().strip('"') for s in m.group(2).split(",")]))
    except Exception as e:
        print(e)
    
    return line.replace(m.group(0), '"{0}"'.format(result))
    
string_function_routine = routine.Routine(
    matcher=string_function_matcher,
    handler=string_function_handler,
    recursive_on_change=True
)

# ---------------------------------------------------------------------------------- #

# Merge strings that need to be concatenated 
def merge_strings_matcher(line):
    lexer = VBLexer()
    expr = exprFactory.get('string.concatenate')
    
    return expr.search(lexer.tokenize(line)) != None
    
def merge_strings_handler(line):
    lexer = VBLexer()
    tokens = lexer.tokenize(line)
    
    expr = exprFactory.get('string.concatenate')
    
    m = expr.search(tokens)
    
    # Sanity check (should be checked in the merge_strings_matcher)
    if m is not None:
        start,end = m.span
    
        s1 = tokens[start]      # This is the 1st string
        s2 = tokens[end - 1]    # This is the 2nd string
    
        return combine(' ', *[
            lexer.untokenize(tokens[:start]),
            '"{0}{1}"'.format(s1.data, s2.data),
            lexer.untokenize(tokens[end:])
        ])
    
    # Something went wrong.. just return the line
    return line

# We want to run this as a post-routine to allow our other deobfuscators resolve
# more complex techniques to reveal our final strings
merge_strings_routine = routine.Routine(
    matcher=merge_strings_matcher,
    handler=merge_strings_handler,
    recursive_on_change=True,
    type=routine.PRE_ROUTINE|routine.POST_ROUTINE
)

# ---------------------------------------------------------------------------------- #

def print_strings_handler(data):    
    lexer = VBLexer()
    tokenized_code = lexer.tokenize(data)
    
    print('----- Strings -----')
    
    for line in tokenized_code:
        for token in line:
            if type(token) == StringToken:
                print(token)
    
print_strings_routine = routine.Routine(
    handler=print_strings_handler,
    type=routine.MULTILINE_ROUTINE,
    skip_on_match=True,
)

# ---------------------------------------------------------------------------------- #
 
def print_functions_handler(data):
    lexer = VBLexer()
    tokenized_code = lexer.tokenize(data)
    
    print('----- Functions -----')
    
    for line in tokenized_code:
        for token in line:
            if type(token) == FunctionToken:
                print(token)
                
print_functions_routine = routine.Routine(
    handler=print_functions_handler,
    type=routine.MULTILINE_ROUTINE,
    skip_on_match=True,
)

# ---------------------------------------------------------------------------------- #

class VBCodeState(object):
    def __init__(self, stack):
        self.stack = VBStack()
        
        self.current = None
        self.previous = None
        
        self.actions = [
            {
                'matcher': 
                'handler': task_function,
                'callbacks': (),
                'require': ('current', 'previous', 'stack',),
            }
        ]
        
    def require(self, action):
        d = {}
        
        for req in action['require']:
            d[req] = getattr(self, req, None)
            
        return d
        
    # TODO: Check for when we need to push a new stack
    def execute(self, line):
        # Iterate over our actions and see what we need to execute
        for action in self.actions:        
            matcher = action.get('matcher')
            
            # Try to run our matcher
            if not matcher or matcher(line):                
                task = action.get('handler')
                
                if inspect.isfunction(task):
                    result = task(self.require(action))
                
                    # Call our callback functions with our result passed in
                    for callback in action['callbacks']:
                        if not inspect.isfunction(callback):
                            callback = getattr(self, callback)
                        
                        # TODO: Use inspect module to get required kwargs
                        # for our callback and online send in those
                        callback(**result)
            

def resolve_variables_handler(data):
    lexer = VBLexer()
    tokenized_code = lexer.tokenize(data)
    
    state = VBCodeState()
    
    new_code = []
    
    for line in tokenized_code:
        vb_stack.execute(line)

        line = resolve_variables(vb_stack, line)
        
        
    
    return data
    
resolve_variables_routine = routine.Routine(
    handler=resolve_variables_handler,
    type=routine.MULTILINE_ROUTINE,
)
    