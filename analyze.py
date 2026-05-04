import pandas as pd

df = pd.read_csv('data/auto-deck-ai-systems-challenge/train.csv')

train = df[df['split'] == 'train']

print('File size stats (KB):')
print((train['file_size_bytes'] / 1024).describe().round(1))

print('\nExtension breakdown:')
print(train['file_extension'].value_counts())

print('\nMetadata completeness:')
print('Has title:', (train['title'] != '-').sum(), '/', len(train))
print('Has company:', (train['company'] != '-').sum(), '/', len(train))
print('Has creation_date:', (train['creation_date'] != '-').sum(), '/', len(train))
print('Has last_modified:', (train['last_modified'] != '-').sum(), '/', len(train))

print('\nFile size distribution:')
sizes = train['file_size_bytes'] / 1024
print('Under 100KB:', (sizes < 100).sum())
print('100KB-500KB:', ((sizes >= 100) & (sizes < 500)).sum())
print('500KB-1MB:', ((sizes >= 500) & (sizes < 1024)).sum())
print('Over 1MB:', (sizes >= 1024).sum())