class MetricDefinition:
  registered_metrics  = []


  
  def __init__(self, name):
    self.name = name
    self.valid_values = []



  def add_value(self, val):
    self.valid_values.append(val)

  def set_values(self, values):
    self.valid_values = values

  def as_tuple(self):
    return self.name, self.valid_values