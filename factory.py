import types

from utils import (
    LazyObject,
)

# Stores a bunch of expressions in a dictionary
# so that they don't have to be created multiple times   
class DataFactory(object):
    def __init__(self, data=None):
        self.dataset = {}
        
        if isinstance(data, dict):
            self.dataset.update(data)
        
    def add(self, name, data):
        if name in self.dataset:
            raise KeyError(
                "Expression name '{0}' is already in the factory".format(name)
            )

        self.dataset[name] = data
        
    def update(self, name, data):
        self.dataset[name] = data
        
    def get(self, name, default=None):    
        if name in self.dataset:                
            return self.dataset[name]
            
        return default
        
    
if __name__ == "__main__":
    print('Import to use this factory')
