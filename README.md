# PV Design Examples

This repository contains Jupyter notebooks and related materials demonstrating photovoltaic (PV) system modeling and design using [`pvlib`](https://github.com/pvlib/pvlib-python).

## 📦 Environment Setup

### 1. Create and activate the conda environment:
```bash
conda create --name pv_design_examples
conda activate pv_design_examples
```

### 2. Install required packages:
```bash
conda install jupyter matplotlib
pip install pvlib
pip install openpyxl
```

> Note: `pvlib` automatically installs its dependencies, including `pandas`, `numpy`, and `scipy`.

### 3. (Optional) Register the environment as a Jupyter kernel:
```bash
python -m ipykernel install --user --name=pv_design_examples --display-name "Python (pv_design_examples)"
```

### 4. Launch Jupyter Notebook:
```bash
jupyter notebook
```

