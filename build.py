import shutil

shutil.rmtree('./dist/data')
shutil.copytree('./data', './dist/data')
