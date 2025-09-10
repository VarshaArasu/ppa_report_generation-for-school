import matplotlib.pyplot as plt
import numpy as np
items = ['Linguistics', 'Problem Solving', 'Focus and Attention', 'Visual Processing', 'Memory']
post_scores = [26.57, 49.08, 61.28, 57.19, 44.33] 
initial_scores = [20.19, 38.12, 46.45, 48.22, 29.80]  
y = np.arange(len(items))  
height = 0.3
fig, ax = plt.subplots(figsize=(9, 5))
plt.gcf().set_facecolor('darkblue')
ax.set_facecolor('darkblue')  
gap = 0.02
bar1 = ax.barh(y - height/2 - gap, initial_scores, height, label='Initial Assessment', color='orange')  
bar2 = ax.barh(y + height/2 + gap, post_scores, height, label='Post Program Assessment', color='green') 
ax.set_title('Skill Growth - Average Skill Scores    Initial vs Post Program Assessment', color='white')
ax.set_yticks(y)
ax.set_yticklabels(items, color='white') 
ax.tick_params(axis='x', colors='white')  
ax.tick_params(axis='y', colors='white')  
ax.legend(facecolor='gray', edgecolor='white', loc='upper right', fontsize=10, framealpha=0.7)  
for i, rect in enumerate(bar1):
    width = rect.get_width()
    ax.text(width, rect.get_y() + rect.get_height() / 2, str(initial_scores[i]), ha='left', va='center', color='white')
for i, rect in enumerate(bar2):
    width = rect.get_width()
    ax.text(width, rect.get_y() + rect.get_height() / 2, str(post_scores[i]), ha='left', va='center', color='white')
plt.tight_layout()
plt.show()
