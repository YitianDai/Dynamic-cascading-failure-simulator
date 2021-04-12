# Dynamic-cascading-failure-simulator
A dynamic cascading failure simulation platform implemented in DIgSILENT PowerFactory via the Python API. It automatically develops cascading mechanisms, simulates sets of failure scenarios and processes results, and also has good scalability such that it can be easily applied to any power system model. 

# Licensing and Citing
We request that publications derived from the use of dynamic cascading failure simulator explicitly acknowledge that fact by citing the following publication:

Y. Dai, M. Noebels, M. Panteli and R. Preece, "Python Scripting for DIgSILENT PowerFactory: Enhancing Dynamic Modelling of Cascading Failures," 2021 IEEE Madrid PowerTech, PowerTech 2021, 2021.

# Getting Started
The following steps will guide you through getting dynamic cascading failure simulator run on your computer.

## Prerequisites
* DIgSILENT PowerFactory 2020 SP1, MATLAB version 9.4 (R2019a) and Python 3.7 or later are recommended.
* Matpower 7.1 or later is required. Please follow the instructions on the [Matpower Website](https://matpower.org/about/get-started/) for installation and test.
* MATLAB Engine API for Python is required, as it provides a package for Python to call MATLAB as a computational engine. Instructions on how to Install MATLAB Engine API for Python can be found on the [MATLAB website](https://uk.mathworks.com/help/matlab/matlab-engine-for-python.html).
* Test the installation of MATLAB Engine API for Python by running the following command in Python:
```
import matlab.engine
eng = matlab.engine.start_matlab()
```
# Usage
## Installation
1. Clone the repository
```
git clone https://github.com/YitianDai/Dynamic-cascading-failure-simulator.git
```
2. Create a Python script in DIgSILENT PowerFactory (Library->Scripts->New Object) and add the location of the dynamic simulator folder to the path.
## Power system model set-up
Use the following example to start PowerFactory, activate projects and study cases, add controllers and protection relays to the related components, and modify parameters. 
```
% import PowerFactory in engine mode
CascadeSim(project_name='*.IntPrj', study_case_name='*.IntCase')

% Add generator controllers and protection relays to the related components
Set_Generator_Control(gen_control_name='*.blk')
Add_Thermal_Relay()
Add_UFLS_Relay()
Add_OFGT_Relay()
```
##	Dynamic simulation of cascading failure
Use the following example to simulate various failure scenarios and record the signals of interest.(Here exemplifies the simulation of N-2 contingencies)
```
Create_Outage_Event(event_name, target, event_start_time)
Run_Dynamic_Simulation(simulation_type, simulation_time)
```
This should export a .txt file containing all the messages in the output window.
## Data processing
Based on the results obtained in the last step, here describes the execution of various data analysis tasks, including identifying cascading propagation paths, calculating demand loss, and data visualization. 
The sample code is as follows:
```
[disconnected_line, disconnected_gen, load_shedding_matrix] = Record_Cascade_Path(output_window_file, load_power, shedding_percent)
Power_loss = Python_Call_MATLAB(disconnected_line, disconnected_gen, load_shedding_matrix):
```
# Extension
This project aims at the joint operation of the three platforms(MATLAB, Python and PowerFactory DIgSILENT), and replaces the operation that would otherwise need to be manually edited with an automated program. Here, Python is the primary control, calling/controlling the operation of MATLAB and POWER from the externally. The initial focus of the example is on large-scale cascading fault analysis, and various extended applications can be developed based on this automation platform, such as:
* Co-simulation of MATLAB and PowerFactory DIgSILENT：controller/control strategy can be designed in MATLAB and synchronous testing can be performed in PowerFactory DIgSILENT.
* Analysis of economic generation dispatch: import the results of MATPOWER into PowerFactory DIgSILENT through this platform.


