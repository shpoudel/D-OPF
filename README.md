# D-OPF
This repository contains a toolbox for Distribution Optimal Power Flow (D-OPF) algorithms developed by Dr. Anamika Dubey's research group (https://eecs.wsu.edu/~adubey/index.html).

## D-OPF Layout

The following is the recommended structure for extracting data from OpenDSS file:

```console
D-OPF
├── D-Net
│   └── network_model.py
│   └── OpenDSS_CDF.py
│   └── Printed JSON files
├── IEEE-123-Bus
│   └── DSS files
│   └── Printed TXT files
├── IEEE-123-Bus-3Phase
│   └── DSS files
│   └── Printed TXT files
├── LICENSE
└── README.md
```

## Execution

The following procedure will give the JSON/TXT files required to model OPF for a distribution network.

1. Clone the D-OPF repository
    ```console
    C:\....\> git clone https://github.com/shpoudel/D-OPF
    C:\....\> cd D-OPF
    C:\....\D-OPF>
    ```
1. Run the network_model.py or OpenDSS_CDF.py
    ```console
    C:\....\D-OPF> cd D-Net
    C:\....\D-OPF\D-Net>
    C:\....\D-OPF\D-Net> python network_model.py file_type
    Note: file_type can be 'JSON' or 'TXT'
    
    OR 
    
    C:\....\D-OPF\D-Net> python OpenDSS_CDF.py
    
    ```
1. Inside the D-Net folder, you will have the required JSON files if you invoke network_model.py. Similarly, inside the IEEE-123-Bus/IEEE123-Bus-3Phase folder, you will have the required TXT files including the network model in form of CDF format if you invoke OpenDSS_CDF.py.
