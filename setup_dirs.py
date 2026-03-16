import os
dirs = [
    'd:/miro/dev/anliku/hotel_case_writer/src',
    'd:/miro/dev/anliku/hotel_case_writer/prompts',
    'd:/miro/dev/anliku/hotel_case_writer/sample_data',
    'd:/miro/dev/anliku/hotel_case_writer/output',
]
for d in dirs:
    os.makedirs(d, exist_ok=True)
print('Directories created.')
