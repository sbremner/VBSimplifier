import types
import hashlib
import inspect


# Utility function to help with bounds checking
def bound_check(ary, index):
    if isinstance(ary, list):
        return 0 <= index < len(ary)
    return False

    
def md5sum(data):
    m = hashlib.md5()    
    
    if isinstance(data, list):
        for d in data:
            m.update(d)
    elif data is not None:
        m.update(data)
        
    return m.hexdigest()

    
def get_padding(data):
    return len(data) - len(data.lstrip())


# Ignores empty strings/data
def combine(ch=' ', *args):
    output = []
    
    for a in args:
        if a:
            output.append(a)
    
    return ch.join(output)
    

# TODO: Implement additional special methods to make the lazy object
# more robust.
# https://docs.python.org/3/reference/datamodel.html#special-method-names
class LazyObject(object):
    def __init__(self, module=None, modulename=None):
        self.module = module
        
        if not self.module:
            self.module = inspect.getmodule(inspect.stack()[1][0])
            
        self.modulename = modulename
        
    @property
    def object(self):
        return self.resolve()
        
    def __repr__(self):
        return self.object.__repr__()
        
    def __str__(self):
        return self.object.__str__()

    def __hash__(self):
        return self.object.__hash__()
      
    def __bool__(self):
        return self.object.__bool__()
      
    def __getattr__(self, key):            
        if hasattr(self.object, key):
            return getattr(self.object, key, None)
            
        return getattr(self, key, None)

    def __eq__(self, other):
        # Resolve other if it is a lazy object
        if isinstance(other, LazyObject):
            other = other.resolve()
        
        # Types and values should be the same
        return type(self.object) == type(other) and self.object == other
        
    def __ne__(self, other):
        return not self.__eq__(other)

    # If our lazy object happens to be a function, we can
    # call it without being force to call 'resolve'
    def __call__(self, *args):
        f = self.resolve()
        
        if not inspect.isfunction(f):
            raise TypeError('LazyObject is not callable: {0}'.format(f))

        return f(*args)
        
    def resolve(self, default=None):
        return getattr(self.module, self.modulename, default)

        
def lazy(modulename):
    mod = inspect.getmodule(inspect.stack()[1][0])
    return LazyObject(module=mod, modulename=modulename)
    