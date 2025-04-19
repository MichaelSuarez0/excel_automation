import time
import importlib
from typing import Tuple

def import_time():
    start = time.time()
    import excel_automation
    end = time.time()

    print(f"Time to import: {end - start:.4f} seconds")


def measure_import_time(module_name: str) -> Tuple[str, float]:
    """
    Dynamically imports a module by name and measures the import time.

    Parameters
    ----------
    module_name : str
        Name of the module to import.

    Returns
    -------
    Tuple[str, float]
        A tuple with the module name and the time taken in seconds.
    """
    start = time.time()
    importlib.import_module(module_name)
    end = time.time()
    
    print(f"Time to import: {end - start:.4f} seconds")

# pyinstrument -r html -o import_excel_automation.html -c "import excel_automation"

if __name__ == "__main__":
    #import_time()
    measure_import_time("pandas")
    measure_import_time("polars")