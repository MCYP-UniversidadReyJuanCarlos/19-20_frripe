# Calculation of Target Security Level (IEC 62443)

## Project Description

This application calculates the target security level of a zone or conduit by performing a quantitative risk analysis of different scenarios.

Risk analysis considers the economic impact of each scenario, based on a Monte Carlo simulation, and the probability of occurrence of that scenario from an assessment of the vulnerabilities of the assets included in the zone or conduit. 

Finally, the target security level is obtained through the comparison of the risk associated with the different scenarios and the company's EBITDA. 



## Features
- The use of the cyber-risk scenario concept as part of the methodology.
- The quantitative estimation of losses from a range of economic impact defined with 90% CI.
- Include an objective estimation of the probability from the taxonomy of vulnerabilities defined in NIST SP 800 82 and the seven domains defined in IEC 62443.
- It proposes a criteria to estimate probability of a scenario based on the assessment of vulnerabilities related to that domain.
- It estimates the risk in a quantitative way (average inherent loss), using Monte Carlo simulation techniques. 
- Determination of the target security level of a zone from the quantification made and its comparison with the EBITDA.
- Provides the possibility of calculating the mitigated cyber risk (probability reduction) after the implementation of countermeasures. 
- Easily document risk analysis.



## How to run 

![ribbon](.\ribbon.png)

- Scenario Probability: opens a form to estimate quickly the probability of its occurrence based on the vulnerabilities associated with the assets that make up the zone or conduit
- Simulation: performs a Monte Carlo simulation of the losses caused in the different scenarios. As a result of the simulation is obtained:
        o Simulated inherent loss of the scenarios included in each domain.
        o Weighted probability of the scenarios assigned to each domain.
        o Target Security Level of the zone or conduit.
        o Simulated residual loss of the scenarios corresponding to each domain (only if compensatory measures are defined)
- Show simulation data: shows seven sheets, one sheet per domains, with all the data obtained in the simulation.
- Get report: generates a report with the complete risk analysis.



## Basic usage

![diagram](.\diagram.png)



