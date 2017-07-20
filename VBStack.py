from lexer import VariableToken

# This is a stack variable
# (contains the variable and its most recently known value)
class VBStackVariable(object):
    def __init__(self, variable, value=None):
        self.variable = variable
        self.data = value

    # This let's us override the "in" operator:
    # e.g.
    #   vb = VBStackVariable()
    #   if key in vb:
    #       pass
    def __contains__(self, key):
        return False
        
    # Override equality operator to check if 2 variables
    # are equal (we care only about the name)
    def __eq__(self, other):
        return isinstance(other, VBStackVariable) and self.name == other.name
    
    # Not equal operator
    def __ne__(self, other):
        return not self.__eq__(self)
    
    # Override the hash operator to allow for set comparisons with 'in'
    def __hash__(self):
        return hash(self.name)
    
    def __str__(self):
        return '{0}({1})'.format(self.variable, self.value)
        
    def __repr__(self):
        return '{0}({1})'.format(self.variable, self.value)
    
    @property
    def name(self):
        return self.variable.__str__()
        
    @property
    def value(self):
        return self.data
        
    def update(self, value):
        self.data = value
        
        
class VBStack(object):
    def __init__(self, push_on_create=True):
        self.stack = []
        
        # Push the new stack automatically when we 
        # initialize our VBStack object (so we can
        # just automatically start adding variables)
        if push_on_create:
            self.push()
        
    # Use this to see if a variable exists in the stack
    def __contains__(self, key):
        return key in self.variables
        
    @property
    def variables(self):
        vars = []
        
        # Combine all of our variable stacks so we can
        # look at this stack as a single scope
        for s in self.stack:
            vars += s
            
        return vars
        
    @property
    def top(self):
        return len(self.stack) - 1
        
    # Creates a new stack (just an empty list for us)
    def push(self):
        self.stack.append([])
    
    # Removes the variables from our top stack and returns
    # any information we had only from that stack (scope)
    def pop(self):
        return self.stack.pop()
        
    # Adds a variable to our current stack (active scope)
    def add(self, variable, unique=True):
        # Enforce unique variable names when we add a variable
        if unique is True and variable in self.variables:
            raise KeyError("VBStack 'unique=true' enforced: {0}".format(variable))
        
        self.stack[self.top].append(variable)
        
    # Updates a variable that is within our stack (can be at any tier since all
    # stack levels are accessible by the active scope)
    def update(self, variable):
        # Variable isn't in our stack... just add it and return True
        if variable not in self.variables:
            self.add(variable)
            return True
        
        # Need to locate and update our variable
        for i in range(0, len(self.stack)):
            for j in range(0, len(self.stack[i])):
                # Use our equality operator to compare our stack variable
                # since stack variable equality is based on the variables
                # name and not the value
                if variable == self.stack[i][j]:
                    # Force the update to our stack variable
                    self.stack[i][j] = variable
                    # Return True since we succeeded
                    return True
        
        # Unable to update
        return False
            
    def get(self, variable, default=None):
        if not isinstance(variable, VBStackVariable):
            variable = VBStackVariable(variable)
            
        for var in self.variables:
            if variable == var:
                return var
        
        # Couldn't find it
        return default
    
    # Try to resolve a variable with information we know in our stack
    def resolve(self, stack_variable):
        # Attempt to get the stack variable:
        stack_variable = self.get(stack_variable, default=stack_variable)
    
        # Don't resolve non-VBStackVariables, just return them
        if not isinstance(stack_variable, VBStackVariable):
            return stack_variable
            
        # Wasn't in the stack, just return the raw value
        if stack_variable not in self.stack:        
            return stack_variable.variable
            
        result = []
        
        for tok in stack_variable.value:
            if isinstance(tok, VariableToken):
                tok = resolve(self.get(tok, default=tok))
            elif isinstance(tok, FunctionToken):
                args = []
                
                for arg in tok.arguments:
                    args.append([resolves(self.get(t, default=t)) for t in arg])
                    
                tok = FunctionToken('{0}({1})'.format(
                    tok.name,       # Function name
                    ','.join(       # Function arguments
                        [' '.join(a) for a in args]
                    )
                ))
                
            # If it isn't a variable token or a function token,
            # it will just be appended blindly (Numeric/String/etc)
            result.append(tok)
    
        return result
    
    # Used to convert a VBStack into a dictionary of 
    # Key : Value pairs for the VBStackVariables
    def dict(self):
        d = {}
        
        for var in self.variables:
            # Use the value as the raw data
            d[var.name] = var.data
            
        return d
        
    def dumps(self):
        # Iterate over our stack in reverse order
        for i in range(0, len(self.stack)):
            print('Stack #{0}:'.format(self.top - i))
            for var in self.stack[self.top - i]:
                print('> {0}'.format(var))
        

def main():
    stack = VBStack(push_on_create=True)
    five = VBStackVariable(VariableToken('num'), value=5)
    
    stack.add(five)
    stack.add(VBStackVariable(VariableToken('num_5'), value=five))
    stack.add(VBStackVariable(VariableToken('var_1'), value="Data Stuff 1"))
    
    stack.push()
    
    stack.add(VBStackVariable(VariableToken('var_2'), value="Data Stuff 2"))
    
    stack.update(VBStackVariable(VariableToken('var_1'), value="Updated var_1"))
    
    stack.dumps()
    
    stack.pop()
    
    stack.dumps()    
    
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print('User sent keyboard interrupt - exiting')