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

P_i_min <= P_i <= P_i_max

---

## 3. OPF Problem Formulation

The OPF problem is formulated as a weighted multi-objective function:

J = w1 * P_loss + w2 * P_slack + w3 * V_dev

Where:

- **P_loss** : Total real power loss in the network (kW)  
- **P_slack** : Real power imported from the slack bus (kW)  
- **V_dev** : Cumulative voltage deviation, calculated as  
  V_dev = Σ |V_k − 1.0|  
- **w1, w2, w3** : Weighting coefficients representing the relative importance of each objective

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

## 5. Software and Tools

- Python  
- OpenDSS  
- Docker  
- NumPy / Pandas  

---

## 6. Installation and Execution

### 6.1 Prerequisites

- Docker Desktop  
- Python 3.x (if running outside container)  

### 6.2 Clone Repository

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

## 7. Key Contributions

- Multi-objective OPF formulation for radial distribution systems  
- Coordinated DER dispatch strategy  
- Slack bus minimization approach  
- Fully containerized execution framework  
- Research-oriented modular architecture  

---

## 8. Conclusion

This project demonstrates the practical implementation of a distribution-level OPF framework for coordinated DER dispatch. The results confirm improved system efficiency, reduced upstream dependency, and enhanced voltage regulation while maintaining operational security.

The Docker-based deployment enables repeatable experimentation and facilitates integration with higher-level distribution automation services.
