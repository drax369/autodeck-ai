import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
import pickle

# Load features
features_df = pd.read_csv('train_features.csv')

print(f"Training on {len(features_df)} files...")

# Create labels based on actual data distribution
# Using percentiles so classes are balanced

# Complexity — based on slide count percentiles
slide_33 = features_df['slide_count'].quantile(0.33)
slide_66 = features_df['slide_count'].quantile(0.66)
print(f"Slide count thresholds: Low<{slide_33:.0f}, Medium<{slide_66:.0f}, High>={slide_66:.0f}")

def label_complexity(slides):
    if slides <= slide_33:
        return "Low"
    elif slides <= slide_66:
        return "Medium"
    else:
        return "High"

# Density — based on avg words per slide percentiles
word_33 = features_df['avg_words_per_slide'].quantile(0.33)
word_66 = features_df['avg_words_per_slide'].quantile(0.66)
print(f"Avg words thresholds: Light<{word_33:.1f}, Medium<{word_66:.1f}, Dense>={word_66:.1f}")

def label_density(avg_words):
    if avg_words <= word_33:
        return "Light"
    elif avg_words <= word_66:
        return "Medium"
    else:
        return "Dense"

# Category from extension
def label_category(ext):
    ext = ext.lower()
    if ext == '.pptx':
        return "Modern"
    elif ext == '.ppt':
        return "Legacy"
    elif ext in ['.pps', '.ppsx']:
        return "Slideshow"
    elif ext in ['.pot', '.potx']:
        return "Template"
    else:
        return "Modern"

# Create labels
features_df['complexity'] = features_df['slide_count'].apply(label_complexity)
features_df['density'] = features_df['avg_words_per_slide'].apply(label_density)
features_df['category'] = features_df['file_extension'].apply(label_category)

print("\nLabel distributions:")
print("Complexity:", features_df['complexity'].value_counts().to_dict())
print("Density:   ", features_df['density'].value_counts().to_dict())
print("Category:  ", features_df['category'].value_counts().to_dict())

# Features for ML model
feature_cols = ['file_size_kb', 'file_size_mb', 'slide_count', 
                'word_count', 'avg_words_per_slide']
X = features_df[feature_cols].values

# Train complexity model
le_complexity = LabelEncoder()
y_complexity = le_complexity.fit_transform(features_df['complexity'])
clf_complexity = RandomForestClassifier(n_estimators=100, random_state=42)
clf_complexity.fit(X, y_complexity)
print(f"\nComplexity model trained. Classes: {le_complexity.classes_}")

# Train density model
le_density = LabelEncoder()
y_density = le_density.fit_transform(features_df['density'])
clf_density = RandomForestClassifier(n_estimators=100, random_state=42)
clf_density.fit(X, y_density)
print(f"Density model trained. Classes: {le_density.classes_}")

# Save everything
models = {
    'clf_complexity': clf_complexity,
    'clf_density': clf_density,
    'le_complexity': le_complexity,
    'le_density': le_density,
    'feature_cols': feature_cols,
    'slide_33': slide_33,
    'slide_66': slide_66,
    'word_33': word_33,
    'word_66': word_66,
}

with open('models.pkl', 'wb') as f:
    pickle.dump(models, f)

print("\nModels saved to models.pkl!")

# Quick accuracy check on training data
complexity_preds = le_complexity.inverse_transform(clf_complexity.predict(X))
density_preds = le_density.inverse_transform(clf_density.predict(X))

complexity_acc = (complexity_preds == features_df['complexity'].values).mean()
density_acc = (density_preds == features_df['density'].values).mean()

print(f"\nTraining accuracy (sanity check):")
print(f"Complexity: {complexity_acc:.1%}")
print(f"Density:    {density_acc:.1%}")