import re

f = open(r'publish\CoverageResults\results.xml', 'r')
lines = f.readlines()

if not re.match('.+Summary.+', lines[2]):
    print(f'Line 2 of results.xml no longer contains the coverage session summary. ' +
          f'Update {__file__}.')
else:
    print(lines[2])
    sequence_coverage = re.findall(r'(?<=sequenceCoverage=")\d+\.\d+', lines[2])[0]
    print(f'Sequence Coverage: {sequence_coverage}')
