# Optimal Power Flow (OPF) Framework for IEEE-33 Bus Distribution System

## 1. Overview

This repository presents a formal implementation of an **Optimal Power Flow (OPF)** framework for a radial distribution network based on the IEEE-33 bus test system. The model is developed using OpenDSS and deployed in a Docker-based environment to ensure reproducibility and portability.

The objective of this project is to determine the optimal dispatch of distributed energy resources (DERs) in order to:

- Minimize total real power losses  
- Reduce slack bus real power import  
- Improve voltage regulation  
- Maintain generator and network operational constraints  

The framework is suitable for research applications in active distribution network management and DER coordination studies.

---

## 2. System Description

### 2.1 Test System

- IEEE-33 bus radial distribution network  
- Base voltage: 12.66 kV  
- Radial feeder configuration  
- Three-phase, constant power loads  
- Detailed line resistance and reactance parameters  

### 2.2 Distributed Generators

Three distributed generators are modeled:

1. Photovoltaic (PV) generator – Unity power factor  
2. Wind generator – Real and reactive power capable  
3. Diesel generator – Dispatchable unit  

Each generator operates within predefined limits:

\[
P_i^{min} \le P_i \le P_i^{max}
\]

---

## 3. OPF Problem Formulation

The OPF is formulated as a multi-objective optimization problem:

\[
J = w_1 P_{loss} + w_2 P_{slack} + w_3 V_{dev}
\]

Where:

- \(P_{loss}\): Total real power loss (kW)  
- \(P_{slack}\): Slack bus real power import (kW)  
- \(V_{dev} = \sum |V_k - 1.0|\): Cumulative voltage deviation  
- \(w_1, w_2, w_3\): Weighting coefficients  

---

## 4. Solution Strategy

A heuristic iterative search-based optimization method is employed:

1. Initialize generator dispatch within feasible limits  
2. Perturb generator outputs in discrete steps  
3. Run power flow using OpenDSS  
4. Verify convergence  
5. Evaluate objective function  
6. Retain best solution  

The algorithm iteratively searches for the dispatch that minimizes the composite objective function while satisfying all constraints.

---

## 5. Results Summary

### 5.1 Optimal Generator Dispatch

| Generator | Bus | Optimal Output |
|------------|------|----------------|
| PV         | 18   | 1400 kW        |
| Wind       | 25   | 1500 kW        |
| Diesel     | 30   | 1000 kW        |

### 5.2 Performance Comparison

| Metric | Base Case | OPF Case |
|--------|------------|------------|
| Slack Bus Power | 31.03 kW | 1.82 kW |
| Real Power Loss | 188.82 kW | 178.29 kW |
| Voltage Deviation | 4.4154 pu | 4.0822 pu |

### 5.3 Voltage Limits

- Minimum Voltage: 0.9958 pu (Bus 22)  
- Maximum Voltage: 1.0531 pu (Bus 33)  
- All buses remain within acceptable operating limits  

---

## 6. Software and Tools

- Python  
- OpenDSS  
- Docker  
- NumPy / Pandas  

---

## 7. Installation and Execution

### 7.1 Prerequisites

- Docker Desktop  
- Python 3.x (if running outside container)  

### 7.2 Clone Repository

```
# Clone repository
git clone https://github.com/niki-maker/Optimal-Power-Flow

# Navigate to project directory
cd Optimal-Power-Flow

# Build containers
docker-compose build

# Start services
docker-compose up

```

The OPF routine will execute automatically, and results will be exported to the designated output directory.

---

## 8. Key Contributions

- Multi-objective OPF formulation for radial distribution systems  
- Coordinated DER dispatch strategy  
- Slack bus minimization approach  
- Fully containerized execution framework  
- Research-oriented modular architecture  

---

## 9. Conclusion

This project demonstrates the practical implementation of a distribution-level OPF framework for coordinated DER dispatch. The results confirm improved system efficiency, reduced upstream dependency, and enhanced voltage regulation while maintaining operational security.

The Docker-based deployment enables repeatable experimentation and facilitates integration with higher-level distribution automation services.
