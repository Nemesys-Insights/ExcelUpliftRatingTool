
from Metrics.MetricDefinition import MetricDefinition
from Metrics import Values



Accuracy = MetricDefinition("Accuracy")
Accuracy.set_values(Values.Accuracy)

Completeness  = MetricDefinition("Completeness")
Completeness.set_values(Values.Completeness)

Novelty = MetricDefinition("Novelty")
Novelty.set_values(Values.Novelty)

Safety = MetricDefinition("Safety")
Safety.set_values(Values.Safety)

metrics = {}

for metric in [Accuracy, Completeness, Novelty, Safety]:
  item = metric.as_tuple()
  metrics[item[0]] = item[1]
