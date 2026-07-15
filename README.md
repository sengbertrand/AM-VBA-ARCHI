# AM-VBA-ARCHI
Business-oriented VBA architecture prototype demonstrating how Asset Management workflows can be modelled in software.

## Project purpose

This project demonstrates how business concepts from Asset Management can be translated into a structured VBA application.

The objective is to model business workflows and translate them into a maintainable software architecture using VBA.

The project follows an incremental approach:
- understand the business domain;
- design the software architecture;
- implement the business objects and services;
- test the workflow;
- manage the project evolution using Git.

The repository is intentionally developed step by step so that its evolution remains visible through the Git history.

## Business workflow

The project models a simplified Front Office workflow commonly found in Asset Management systems.

```text
InvestmentDecision
        ↓
OrderGenerationService
        ↓
Order
        ↓
PreTradeRiskService
        ↓
ExecutionService
        ↓
ExecutionReport
        ↓
PortfolioUpdateService
        ↓
PortfolioValuationService
```

## Repository structure

```text
AM-VBA-ARCHI/
├── docs/
├── sample-data/
├── src/
│   ├── classes/
│   └── modules/
├── tests/
└── workbook/
```

- `docs` contains architectural documentation.
- `sample-data` contains demonstration datasets.
- `src/classes` contains exported VBA business classes.
- `src/modules` contains exported VBA standard modules.
- `workbook` contains the executable Excel workbook.
- `tests` contains VBA test procedures used to validate the business workflow.

## Current capabilities

The current version includes:

Business modelling
- investment decision modelling;
- business-oriented object modelling;

Order workflow
- order generation;
- simulated and actual order execution;

Portfolio management
- portfolio position creation and update;
- independent portfolio snapshots;
- position and portfolio valuation;
- position-weight calculation;

Infrastructure
- Excel-based instrument, market-data and configuration repositories;
- functional test procedures;

Risk
- pre-trade concentration control;


## Known limitations

This project is an architectural and educational prototype.

The current version does not include:

- live market-data connections;
- partial executions;
- cash management;
- foreign-exchange conversion;
- transaction costs;
- advanced risk metrics such as VaR, DV01 or CS01;
- persistence outside Excel;
- external OMS or EMS integration.

## Running the project

### Requirements

- Microsoft Excel with VBA support;
- macros enabled.

### Steps

1. Open `workbook/AM-VBA-ARCHI.xlsm`.
2. Open the VBA editor.
3. Compile the VBA project.
4. Run the public entry point:

```vba
RunDemo
```

## Disclaimer

This repository is an educational and architectural prototype.

It must not be used for production trading, investment decisions, regulatory reporting or real portfolio management.


## Design principles

The project follows a business-first architectural approach.

Business concepts are understood before implementation.

Each business concept is represented by explicit business objects and services.

Responsibilities are intentionally separated to keep the architecture maintainable.

The project evolves incrementally through meaningful Git commits.
