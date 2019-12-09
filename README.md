# D-OPF
This repository contains a toolbox for Distribution Optimal Power Flow (D-OPF) Algorithms developed by Dr. Anamika Dubey's research group (https://eecs.wsu.edu/~adubey/index.html).

## D-OPF Layout

The following is the recommended structure for extracting data from OpenDSS file:

```console
.
├── D-NET
│   └── network_model.py
│   └── Printed JSON files
├── IEEE-123-Bus
│   └── DSS files
├── LICENSE
└── README.md
```

## Execution

The following procedure will give the JSON files required to moel OPF for a distribution network

1. Clone the D-OPF repository
    ```console
    git clone https://github.com/shpoudel/D-OPF
    cd D-OPF
    ```
1. Run the network_model.py
    ```console
    cd D-NET
    python network_model.py
    ```
1. Once inside the D-Net folder, you will have the required JSON files
