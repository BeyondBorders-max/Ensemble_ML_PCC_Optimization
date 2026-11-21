# Ensemble Machine Learning for Multi-Objective Optimization of Amine-Based Post-Combustion Carbon Capture with Shap and Counterfactual Explanation

This repository contains the code and data associated with the extended study on multi-objective optimization of amine-based post-combustion carbon capture (PCC) processes. Building on Hanifi et al. (2026), this work substantially expands the modeling and optimization framework with a richer dataset, a diverse ensemble of surrogate models, six simultaneous optimization targets, and a Knee-Point decision-making strategy for Pareto solution selection and dual XAI.

## Overview

Post-combustion carbon capture (PCC) using amine-based absorption remains one of the most viable near-term technologies for reducing CO₂ emissions from existing fossil-fuel power plants. This work presents an integrated many-objective framework combining:

- **Expanded Rigorous Process Simulation:** A dataset of **5,200 simulation rows** spanning five amine solvents (MEA, DGA, DEA, TEA, MDEA) generated via a Python–Aspen Plus interface using the ENRTL-RK thermodynamic model. Starting from **34 initial input variables**, feature selection via Multi-Random Forest Gini importance (threshold < 0.005) reduces the working set to **30 final inputs**.
- **Eight Surrogate Models:** ANN, TabPFN, PSO-CatBoost, PSO-AdaBoost, PSO-XGBoost, PSO-GBR, Bagged CatBoost (6× base learners), and a Stacked Meta-Surrogate — covering individual, PSO-tuned, bagged, and stacked architectures.
- **Six Optimization Targets:** CO₂ recovery (%), CO₂ purity (%), specific energy consumption — SEC (kJ/kg CO₂), solvent loss (kg/h), OPEX ($/t CO₂), and corrosion rate (mm/yr).
- **Multi-Objective Optimization:** The best-performing surrogate is coupled with **NSGA-III** to explore high-dimensional trade-offs across all six objectives simultaneously.
- **Aspen Plus Validation & Knee-Point Selection:** 132 Pareto optimal solutions are validated in Aspen Plus, and the final operating point is selected via **Knee-Point** analysis (not TOPSIS).
- **Shap and Counterfactual Explainations ** shap identifies the most influential variables and their effects on performance while DiCE identify the minimum changes in operating conditions required to achieve desired performance targets.
- 
- ### Key Findings

| Model | R² (avg. across targets) |
|---|---|
| Stacked Meta-Surrogate | **0.9956** |
| Bagged CatBoost (6×) | 0.9947 |
| PSO-CatBoost | 0.9942 |

- **132 Pareto solutions** validated in Aspen Plus with an average error < **5.12%** across all targets.
- **Knee-Point optimal solution:**
  - CO₂ recovery: **98.92%**
  - CO₂ purity: **99.45%**
  - SEC: **10,248.5 kJ/kg CO₂**
  - Solvent loss: **21.84 kg/h**
  - OPEX: **$0.89/t CO₂**
  - Corrosion rate: **1.42 mm/yr**
- **DGA (42.6 wt%)** emerged as the optimal solvent for the many-objective formulation.

## Repository Structure

```
PCC-Ensemble-Optimization-main/
├── AsPy_34x6.ipynb               ← Aspen Plus automation (34 inputs, 6 targets)
├── aspen_automation_30x6.py      ← Full 30-variable dataset generator
├── CodeLibrary.py                ← Aspen Plus COM interface
├── Dataset.csv                   ← 5,200 × 36 dataset
├── LICENSE                       ← MIT License
├── PCC_Ensemble_ML.ipynb         ← 8-model surrogate training pipeline notebook
├── PCC_Ensemble_Optimization.ipynb ← NSGA-III many-objective optimization notebook
├── pareto_front_optimal_solutions.csv ← Raw NSGA-III Pareto optimal solutions
├── aspen_validated_pareto_solutions.csv ← 132 Aspen-validated Pareto solutions
├── kneepoint_optimal_solution.csv ← Selected optimal Knee-Point operating conditions
├── scaler_X.pkl / scaler_y.pkl   ← RobustScaler feature and target preprocessors
├── README.md                     ← This file
└── requirements.txt              ← Python dependencies
```

## Pipeline Steps

1. **Aspen Plus simulation data collection** (`AsPy_34x6.ipynb`): Automate Aspen Plus via COM interface to collect 5,200 simulation rows for 5 amines across 34 input variables and 6 output targets.
2. **Dataset assembly** (`aspen_automation_30x6.py`): Aggregate raw simulation outputs into a consolidated `Dataset.csv` (5,200 × 36).
3. **Exploratory data analysis & correlation analysis**: Visualize distributions, pairwise correlations, and amine-wise statistics.
4. **Data preprocessing**: Standardization/normalization, train/test split (e.g., 80/20), and outlier handling.
5. **Feature selection — Multi-Random Forest** (Gini importance < 0.005 threshold removes 4 variables → **30 final inputs**): Rank input importance across multiple Random Forest runs; drop features that consistently fall below the Gini threshold.
6. **Surrogate model training**: Train 6 individual models (ANN, TabPFN, PSO-CatBoost, PSO-AdaBoost, PSO-XGBoost, PSO-GBR) plus Bagged CatBoost (6 base learners) and a Stacked Meta-Surrogate on the 30-input dataset.
7. **Model evaluation & cross-validation**: 5×5 repeated cross-validation, R², RMSE, and MAE reporting for all 8 models across all 6 targets.
8. **NSGA-III many-objective optimization** (`PCC_Ensemble_Optimization.ipynb`): Use the best surrogate (Stacked Meta-Surrogate) as a fitness function inside NSGA-III (pymoo) with 6 objectives and appropriate bounds.
9. **Aspen Plus validation + Knee-Point selection**: Validate 132 Pareto front solutions in Aspen Plus; apply Knee-Point analysis to identify the single best balanced operating condition.

## Requirements

| Package | Minimum Version |
|---|---|
| Python | 3.10+ |
| TensorFlow | ≥ 2.15 |
| scikit-learn | ≥ 1.3 |
| CatBoost | ≥ 1.2 |
| XGBoost | ≥ 2.0 |
| pymoo | ≥ 0.6 |
| pandas | ≥ 2.0 |
| numpy | ≥ 1.24 |
| scipy | ≥ 1.11 |
| matplotlib | ≥ 3.7 |
| seaborn | ≥ 0.12 |
| joblib | ≥ 1.3 |
| python-docx | ≥ 0.8 |
| pywin32 | — (Aspen COM interface) |

Install all dependencies via:

```bash
pip install -r requirements.txt
```

> **Note:** `pywin32` is required only on Windows for the Aspen Plus COM automation interface (`CodeLibrary.py`). Surrogate training and optimization notebooks can be run without it.

## Citation

If you use this code, data, or methodology in your research, please cite our paper:

```bibtex
@article{Garba2026,
  title   = {Ensemble Machine Learning for Multi-Objective Optimization of Amine-Based Post-Combustion Carbon Capture with shap and Counterfactual explanations},
  journal = {[Energy]},
  year    = {2026},
  author  = {[Garba Muhammad Anchau, Oluwanifemi Boluwatife Ogunsina, Muhammad Malah Bukar, Halimatu Musa, Shehu Abdullahi]},
  keywords = {Amine based absorption; Ensemble Machine Learning; Explainable artificial intelligence; Knee-point;  Multi-Objective Optimization, NSGA-III, Post Combustion Carbon capture}
}
```

If you also use or build on the base methodology from Hanifi et al. (2026), please additionally cite:

```bibtex
@article{HANIFI2026141133,
  title    = {Multi-objective optimization of the amine-based natural gas-fired post-combustion carbon capture process using machine learning surrogate modeling},
  journal  = {Energy},
  volume   = {355},
  pages    = {141133},
  year     = {2026},
  doi      = {https://doi.org/10.1016/j.energy.2026.141133},
  author   = {Kiarash Hanifi and Mohammad Rahmani and Mostafa Haghighi},
  keywords = {Post-combustion carbon capture, Amine-based absorption, Machine learning, ANN, TabPFN, Multi-objective optimization, NSGA-II}
}
```

## Links

- **Repository:** [[(https://github.com/BeyondBorders-max/Ensemble_ML_PCC_Optimization/tree/main/Ensemble_ML_PCC_Optimization]]
- **Base Paper (Hanifi et al. 2026):** [Multi-objective optimization of the amine-based natural gas-fired post-combustion carbon capture process using machine learning surrogate modeling](https://doi.org/10.1016/j.energy.2026.141133)
- **License:** [MIT License](LICENSE)
