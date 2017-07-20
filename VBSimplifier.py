import re
import sys
import hashlib

from optparse import OptionParser, OptionGroup

# Need routine for the constants (Use capital R to avoid collisions)
import routine as RoutineModule

from lexer import (
    VBLexer
)

from utils import (
    bound_check, md5sum, get_padding, combine
)

# Import all of our routines, we can access with routines.name
from routines import (
    routineFactory
)
    
# Static value for INVALID iterations
INVALID_ITERATION_ID = -1
    

class Iteration(object):
    lexer = VBLexer()
    
    def __init__(self, data, from_id=INVALID_ITERATION_ID, to_id=INVALID_ITERATION_ID):        
        # Shortcut version is buggy for some reason? Not sure why
        # self.data = [self.lexer.tokenize(line) for line in data]
        self.data = []
        
        if data is not None:
            for line in data:
                self.data.append(self.lexer.tokenize(line))
        
        self.from_id = from_id
        
        if to_id == INVALID_ITERATION_ID:
            to_id = []
        
        if not isinstance(to_id, list):
            to_id = [to_id]
            
        self.to_id = to_id

    @property
    def lines_of_code(self):
        return len(self.data)

    @property
    def code(self):
        # lexer.dumps returns the code as a list
        # use '\n'.join(lexer.dumps(data)) to combine
        return self.lexer.dumps(self.data)
        
        
    def get_functions(self):
        return index_functions(self.data)
        
    def get_line(self, line_no):
        if bound_check(self.data, line_no):
            return ' '.join([tok.__str__() for tok in self.data[line_no]])
            
        return None       
        
    def branch(self, to_id):
        self.to_id.append(to_id)
        
    def get_branches(self):
        if len(self.to_id) > 0:
            return self.to_id
            
        return None
        

class VBCode(object):
    def __init__(self, file=None, data=None):
        # iterations will contain forward/backward iterations
        # as each routine gets run (this can be used for reversing)
        self.iterations = []
        self.active_iteration = -1

        # This is the file name for our original code
        self.file = file
        
        if file is None:
            # I guess we need to load from data
            self.push(data)
        else:
            self.load_from_file(self.file)
            
    @property
    def current(self):
        if bound_check(self.iterations, self.active_iteration):
            return self.iterations[self.active_iteration]
            
        return None
        
    def load_from_file(self, file):
        with open(file, "r") as f:
            data = f.readlines()
            
        self.push(data)
        
    def get_code(self, iteration=-1):
        if bound_check(self.iterations, iteration):
            return self.iterations[iteration]
            
        return self.current

    def peek(self):
        if bound_check(self.iterations, self.active_iteration):
            return self.current.get_branches()
            
        return None

    # Jump back to a previous iteration
    def back(self, count=1):
        if count >= 1 and bound_check(self.iterations, current.from_id):
            self.jump(current.from_id)
            return self.back(count - 1)
            
        # Count = 0 here (we no longer go back - just return the current iteration
        return self.current
        
    def jump(self, iteration):
        if bound_check(self.iterations, iteration):
            self.active_iteration = iteration
            
        return self.get_code()

    def push(self, data, allow_none=False):
        if data is None and allow_none is False:
            raise ValueError("Code push has no data (use allow_none=True to push empty data)")
    
        # Checks the hex-digest of the current iteration to see if anything changes
        # and will not update if nothing has occurred to the output
        if md5sum(getattr(self.current, 'code', None)) == md5sum(data):
            return self.current
    
        # Build the new item
        next = Iteration(data, from_id=self.active_iteration)
        self.iterations.append(next)
        
        to_id = len(self.iterations) - 1
        
        # Update to "to_id" for our old iteration to point to the new one
        # but only do this if we had an old iteration
        if bound_check(self.iterations, self.active_iteration):
            self.iterations[self.active_iteration].branch(to_id)

        # Update our active iteration tracker
        self.active_iteration = to_id
        
        # Returns the current item (added from this push)
        return self.current


class VBSimplifier(object):
    def __init__(self, file):
        self.vb_code = VBCode(file)

        self.routines = []   
        
    def register(self, routine, execute_on_register=False):
        self.routines.append(routine)
        
        # Check if we should execute this routine as soon as we are registered
        if execute_on_register == True:
            self.run_routine(routine)

    def unregister(self, routine):
        if routine in self.routines:
            self.routines.remove(routine)
            
    def run_routine(self, routine):
        # We handle multi-line routines slightly differently
        # If a routine is multi-line, then it cannot run single line routines
        if (routine.type & RoutineModule.MULTILINE_ROUTINE) == RoutineModule.MULTILINE_ROUTINE:
            return self.run_multiline_routine(routine)
    
        # Keep our old code stored locally to index
        old_code = self.vb_code.get_code()
        # Add all our new code to the temp storage we will push later
        new_code = []
        
        for i in range(0, old_code.lines_of_code):
            line = old_code.get_line(i)

            # Run the routine and save the results
            #   matched: contains whether or not the routine ran the handler
            matched, line = routine.run(line)
            
            # Since we matched, we should check if we are skipping this
            # instead of appending it (useful for stripping useless data
            # such as comments in code)
            if routine.skip_on_match and matched:
                continue
            
            new_code.append(line)
            
        # push() returns the most recently pushed item
        return self.vb_code.push(new_code)

    def run_multiline_routine(self, routine):
        old_code = self.vb_code.current.code
        
        matched, new_code = routine.run(old_code)
        
        if routine.skip_on_match and matched:
            return self.vb_code.current

        # Push the new code since we aren't skipping
        return self.vb_code.push(new_code)
        
    # This will run all of the routines (pre -> routine -> post) unless
    # The final flag to run is MULTILINE
    def run(self, flag=RoutineModule.ALL_ROUTINES):
        for i in range(0, len(bin(flag)[2:])):
            run_flag = 1 << i
            
            if (flag & run_flag) == run_flag:
                for r in self.routines:
                    if (r.type & run_flag) == run_flag:
                        self.run_routine(r)

    def dump(self, file, show_file_no=False):
        # Create a local copy for dumping
        code = self.vb_code.current.code
        
        with open(file, "w") as f:
            for i in range(0, len(code)):
                if show_file_no:
                    f.write('{0} | '.format(i,))
                f.write('{0}\n'.format(code[i]))

# ------------------------------------------------------------------------- #

# This is to modify the behaviour of the OptParser to make it check arguments
# more strictly
class NonCorrectingOptionParser(OptionParser):

    def _match_long_opt(self, opt):
        # Is there an exact match?
        if opt in self._long_opt:
            return opt
        else:
            self.error('"{0}" is not a valid command line option.'.format(opt))

# This method return the parser to parse our command line arguments.

USAGE_MESSAGE = "Type 'python user_history.py --help' for usage."

def get_parser():

    parser = NonCorrectingOptionParser(add_help_option=False)

    # Default 'Options' group (top level)
    parser.add_option('-h', '--help', help='Show help message',
                      action='store_true')
    
    # Deobfuscation Group
    deobfuscation_group = OptionGroup(parser, 'Deobfuscation Options',
    'Toggle deobfuscation tactics below')
    deobfuscation_group.add_option('--all', help='Enable all deobfuscation techniques', action='store_true')

    deobfuscation_group.add_option('--concatenate', help='Concatenates strings as needed', action='store_true')
    deobfuscation_group.add_option('--math', help='Solves math equations', action='store_true')
    deobfuscation_group.add_option('--comments', help='Strips single line comments', action='store_true')
    deobfuscation_group.add_option('--str-functions', help='Resolves common string functions (Left, Right, StrReverse, etc)',
                        action='store_true')
    parser.add_option_group(deobfuscation_group)

    # Input Group
    input_group = OptionGroup(parser, 'Input Options', 'Provide the following data as input to the program')
    input_group.add_option('-i', '--input', help='Input VB Script file to deobfuscate. Required',
                           action='store')
    parser.add_option_group(input_group)

    # Output Group
    output_group = OptionGroup(parser, 'Output Options', 'Specify the output for the program')
    output_group.add_option('-o', '--output', help='Output file to write deobfuscated code', action='store')
    output_group.add_option('-s', '--strings', help='Print strings within code', action='store_true')
    output_group.add_option('-f', '--functions', help='Print functions within code', action='store_true')

    parser.add_option_group(output_group)

    return parser


def print_help(parser):
    print(parser.format_help().strip())
    

def main(opts):
    try:
        vbs = VBSimplifier(opts.input)
        
        if getattr(opts, 'comments', False) or getattr(opts, 'all', False):
            vbs.register(routineFactory.get('comments'))
        
        if getattr(opts, 'math', False) or getattr(opts, 'all', False):
            vbs.register(routineFactory.get('math'))
        
        if getattr(opts, 'str_functions', False) or getattr(opts, 'all', False):
            vbs.register(routineFactory.get('str_functions'))
            
        if getattr(opts, 'concatenate', False) or getattr(opts, 'all', False):
            vbs.register(routineFactory.get('concatenate'))
        
        # vbs.register(routineFactory.get('resolve'))
        
        vbs.run()

        # Print strings/functions if we were asked to
        if getattr(opts, 'strings', False):
            vbs.register(routineFactory.get('strings'))
        if getattr(opts, 'functions', False):
            vbs.register(routineFactory.get('functions'))
        
        if getattr(opts, 'strings', False) or getattr(opts,'functions', False):
            vbs.run(flag=RoutineModule.MULTILINE_ROUTINE)
        
        if getattr(opts, 'output', None):
            vbs.dump(getattr(opts, 'output', None))
    
    except Exception as e:
        print('Unknown error: {0}'.format(e))
    
def is_valid_opts(opts):
    mandatories = ['input']
    
    for m in mandatories:
        if not opts.__dict__[m]:
            print('Missing required argument: {0}'.format(m))
            return False
            
    return True
    
if __name__ == "__main__":
    try:
        parser = get_parser()
        
        opts = parser.parse_args()[0]
        
        if opts.help:
            print_help(parser)
        elif not sys.argv[1:] or not is_valid_opts(opts):
            print(USAGE_MESSAGE + '\n')
        else:
            main(opts)
    except KeyboardInterrupt:
        print('User sent keyboard interrupt - exiting')