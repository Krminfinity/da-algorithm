# Deferred Acceptance (DA) Algorithm Implementation

## Overview

This project implements the Deferred Acceptance (DA) algorithm, also known as the Gale-Shapley algorithm, for solving stable matching problems. The implementation is designed to handle multiple acceptances and uses Excel files for input and output, making it user-friendly for those familiar with spreadsheet software.

## Features

- Implements the Deferred Acceptance algorithm for stable matching
- Supports multiple acceptances (e.g., schools accepting multiple students)
- Reads preferences and capacities from Excel files
- Outputs matching results to an Excel file
- Provides clear, commented code for easy understanding and modification

## Requirements

- Python 3.6 or higher
- openpyxl library

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/Krminfinity/da-algorithm.git
   cd da-algorithm
   ```

2. Install the required library:
   ```
   pip install openpyxl
   ```

## Usage

1. Prepare an Excel file named `matching_data.xlsx` with the following sheets:
   - `Proposers Preferences`: List of proposers and their preferences
   - `Acceptors Preferences`: List of acceptors and their preferences
   - `Acceptors Capacities`: List of acceptors and their capacities

2. Run the script:
   ```
   python da_algorithm.py
   ```

3. The results will be displayed in the console and saved in `matching_results.xlsx`.

## Input File Format

The `matching_data.xlsx` file should be structured as follows:

1. `Proposers Preferences` sheet:
   ```
   Proposer | 1st Choice | 2nd Choice | 3rd Choice | ...
   A        | X          | Y          | Z          | ...
   B        | Z          | Y          | X          | ...
   C        | Y          | Z          | X          | ...
   ```

2. `Acceptors Preferences` sheet:
   ```
   Acceptor | 1st Choice | 2nd Choice | 3rd Choice | ...
   X        | B          | A          | C          | ...
   Y        | C          | B          | A          | ...
   Z        | A          | C          | B          | ...
   ```

3. `Acceptors Capacities` sheet:
   ```
   Acceptor | Capacity
   X        | 2
   Y        | 1
   Z        | 2
   ```

## Output

The script generates an Excel file named `matching_results.xlsx` with the following format:

```
Acceptor | Matched Proposers
X        | A, B
Y        | C
Z        | D, E
```

## Contributing

Contributions to improve the algorithm or extend its functionality are welcome. Please feel free to submit pull requests or open issues to discuss potential changes.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- This implementation is based on the Deferred Acceptance algorithm described by David Gale and Lloyd Shapley in their 1962 paper "College Admissions and the Stability of Marriage".

## Contact

If you have any questions or feedback, please open an issue on this GitHub repository.
