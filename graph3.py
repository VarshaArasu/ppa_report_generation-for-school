import matplotlib.pyplot as plt
import numpy as np
def rgb_to_normalized(rgb):
    return tuple([x / 255.0 for x in rgb])
rgb = (146, 208, 80)  
normalized_rgb = rgb_to_normalized(rgb)
categories = ['Contextual Recall', 'Associative Recall', 'Working Memory',  'Spatial Awareness', 'Conservation', 'Creative Thinking-Visualisation','Creative Thinking-Synthesis','Sustaining attention', 'Selective attention', 'Divided attention', 'Inductive Reasoning', 'Deductive Reasoning', 'Classification', 'Distinguishing', 'Pattern Recognition/Sequencing', 'Inferring', 'Prediction and Conclusion','Critical Thinking-Visualisation', 'Language Processing', 'Assimilation', 'Accomodation']
scores = [52.09, 55.16, 40.98, 62.82, 52.34, 62.03, 51.21, 49.37, 56.95, 68.96, 
          53.43, 40.83, 67.56, 42.11, 47.18, 32.48, 42.71, 63.8, 26.15, 34.58, 66.61]
graph_background_rgb = (32, 56, 100)  
axes_background_rgb = (32, 56, 100)  
graph_background = rgb_to_normalized(graph_background_rgb)
axes_background = rgb_to_normalized(axes_background_rgb)
plt.figure(figsize=(9, 6)) 
plt.gcf().set_facecolor(graph_background)  
ax = plt.gca()
ax.set_facecolor(axes_background)  
bars = plt.bar(categories, scores, color=normalized_rgb)
plt.title('Sub Skill Average Scores', fontsize=16, fontweight='bold', color='white')
plt.xticks(rotation=45, ha='right', fontsize=10, color='white') 
plt.yticks(np.arange(0, 90, 10), color='white')
ax.grid(True, which='both', axis='y', linestyle='-', linewidth=0.5, color='black')
for i, bar in enumerate(bars):
    height = bar.get_height()  
    plt.text(bar.get_x() + bar.get_width() / 2, height + 1, f'{height:.2f}', ha='center', va='bottom', fontsize=10, color='white')
plt.tight_layout()
plt.show()


