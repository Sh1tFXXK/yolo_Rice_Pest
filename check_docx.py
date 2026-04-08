import importlib

mod = importlib.import_module("docx")
print(mod.__version__)
