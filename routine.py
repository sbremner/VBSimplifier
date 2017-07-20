import utils

# Binary flags specify how to run these routines (in this order):
# PRE     = 0000 0001 = 0x01
# ROUTINE = 0000 0010 = 0x02
# POST    = 0000 0100 = 0x04
# MULTI   = 0000 1000 = 0x08
PRE_ROUTINE = 0x1
ROUTINE = 0x2
POST_ROUTINE = 0x4
MULTILINE_ROUTINE = 0x8

# ALL_ROUTINES = 0x0F
ALL_ROUTINES = (PRE_ROUTINE | ROUTINE | POST_ROUTINE | MULTILINE_ROUTINE)
    
# This class is used to specify a matcher/handler routine
# which is used by the VBSimplifier class when deobfuscating
class Routine(object):
    def __init__(self, matcher=None, handler=None, type=ROUTINE, skip_on_match=False, recursive_on_change=False):
        self.matcher = matcher
        self.handler = handler
        self.type = type
        
        # NOTE: This will make the VBSimplifier will not include the results if
        # this matcher is matched (the handler will effectively do nothing except
        # for logging purposes)
        self.skip_on_match = skip_on_match
        
        # NOTE: This will make the VBSimplifier re-run this routine multiple times
        # until the code is no longer changing. This can be useful for nested obfuscation
        # that requires multiple passes to succeed.
        self.recursive_on_change = recursive_on_change
        
    def match(self, line):
        try:
            active_matcher = self.matcher
            
            # Default matcher is to return True
            if not self.matcher:
                return True
        
            # This allows us to chain matchers one after another
            # and requiring all of them to return True
            if isinstance(self.handler, (list, tuple)):                
                for f in self.handler:
                    if not f(line):
                        return False
                
                # None of our handlers return False
                return True
        
            return self.matcher(line)
        except Exception as e:
            print('Error inside matcher: {0}'.format(active_matcher.__name__))
            print('> Message: {0}'.format(e))
            return False
           
    def execute(self, line):
        try:
            active_handler = self.handler
        
            # Default handler is to do nothing
            if not self.handler:
                return line
        
            # This allows us to chain handlers one after another
            # and pass the output from 1 to the next
            if isinstance(self.handler, (list, tuple)):                
                result = line
                
                for f in self.handler:
                    active_handler = f                    
                    result = f(result)
                
                return result
                    
            return self.handler(line)
        except Exception as e:
            print('Error inside handler: {0}'.format(active_handler.__name__))
            print('> Message: {0}'.format(e))
            return line
           
    def run(self, data):
        # Initialize our output with our input data
        output = data
        
        # Check if we match (we need to know this later)
        # If we have no matcher, we just skip this matching step
        matched = (not self.matcher) or self.match(data)
        
        if matched:
            output = self.execute(data)
            
            if utils.md5sum(output) != utils.md5sum(data) and self.recursive_on_change:
                _, output = self.run(output)
        
        # The recursion will handle our output and our matched variable
        # will contain the match if we ever matched (not when the final case fails)
        return matched, output
        
if __name__ == "__main__":
    print('import routine to use this file')