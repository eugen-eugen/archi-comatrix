# Archi Model Metamodel for Comatrix & AppList Scripts

This document describes which elements, relationships, properties, and conventions in your Archi model influence the outputs of the `comatrix` and `applist` scripts.

## Table of Contents
- [Overview](#overview)
- [Element Types](#element-types)
- [Relationship Types](#relationship-types)
- [Specializations](#specializations)
- [Properties](#properties)
- [Naming Conventions](#naming-conventions)
- [Model Structure Requirements](#model-structure-requirements)
- [Impact on Outputs](#impact-on-outputs)

---

## Overview

The comatrix and applist scripts analyze specific patterns in your Archi model to generate Excel reports. Understanding what these scripts look for helps you structure your model to produce accurate and meaningful outputs.

```
┌─────────────────────────────────────────────────────────────┐
│                    Your Archi Model                         │
│                                                             │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐ │
│  │   Grouping   │    │ Application  │    │ Triggering   │ │
│  │ (Fachbereich)│    │  Components  │    │Relationships │ │
│  │              │    │              │    │  (NST_*)     │ │
│  └──────────────┘    └──────────────┘    └──────────────┘ │
│         │                    │                    │         │
│         └────────────────────┴────────────────────┘         │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│              Comatrix & AppList Scripts                     │
└─────────────────────────────────────────────────────────────┘
                           │
           ┌───────────────┴───────────────┐
           ▼                               ▼
┌────────────────────┐         ┌────────────────────┐
│   comatrix.xlsx    │         │   applist.xlsx     │
│ Connectivity Matrix│         │Application Catalog │
└────────────────────┘         └────────────────────┘
```

---

## Element Types

### 1. Application Component (`application-component`)

**Used by:** Both comatrix and applist

**Purpose:** Represents software applications in your architecture.

**What matters:**
- **Name**: Used as the primary identifier in outputs
- **Specialization**: Filters which applications appear in applist (see [Specializations](#specializations))
- **Relationships**: 
  - As source/target of triggering relationships (comatrix)
  - As child in aggregation/composition relationships (domain/Fachbereich assignment)

**Example:**
```
Application Component: "Customer Portal"
  Specialization: "Geschäftsanwendung"
  Grouped by: Domain "Sales"
  Grouped by: Fachbereich "Customer Management"
```

### 2. Grouping (`grouping`)

**Used by:** Both comatrix and applist (indirectly through domain/Fachbereich detection)

**Purpose:** Organizes elements into logical groups (domains, Fachbereich).

**What matters:**
- **Specialization**: Must be "Domäne" or "Fachbereich" to be recognized
- **Name**: Used as the domain/Fachbereich name in outputs
- **Relationships**: Aggregation or composition relationships TO application components

**Example:**
```
Grouping: "Sales Domain"
  Specialization: "Domäne"
  Contains (via aggregation): "Customer Portal", "Order System"
```

---

## Relationship Types

### 1. Triggering Relationship (`triggering-relationship`)

**Used by:** comatrix only

**Purpose:** Represents interactions between application components via interfaces.

**What matters:**
- **Name**: MUST start with `NST_` prefix to be included
- **Source**: The application component that triggers (B-element in matrix)
- **Target**: The application component that is triggered (A-element in matrix)
- **Property "Schnittstelle"**: Interface name (can have multiple values)

**Example:**
```
Triggering Relationship: "NST_OrderPlacement"
  Source: "Customer Portal" (B-element)
  Target: "Order System" (A-element)
  Property "Schnittstelle": "REST API Orders"
```

**Output in comatrix.xlsx:**
- Row: Order System | REST API Orders
- Column: Customer Portal
- Cell: "x"

### 2. Aggregation Relationship (`aggregation-relationship`)

**Used by:** Both scripts (for domain/Fachbereich detection)

**Purpose:** Groups application components into domains or Fachbereich.

**What matters:**
- **Source**: Must be a Grouping with specialization "Domäne" or "Fachbereich"
- **Target**: Application component(s) to be grouped
- **Direction**: Always FROM grouping TO application component

**Example:**
```
Aggregation Relationship
  Source: Grouping "Sales Domain" (specialization: "Domäne")
  Target: Application Component "Customer Portal"
```

**Result:** Customer Portal's domain = "Sales Domain"

### 3. Composition Relationship (`composition-relationship`)

**Used by:** Both scripts (for domain/Fachbereich detection)

**Purpose:** Same as aggregation, but with stronger ownership semantics.

**What matters:** Same as aggregation relationship

**Note:** Both aggregation and composition are treated identically for domain/Fachbereich detection.

---

## Specializations

Specializations are custom labels you can assign to ArchiMate elements. The scripts rely heavily on specific specialization values.

### For Grouping Elements

| Specialization | Used By | Purpose | Example |
|----------------|---------|---------|---------|
| `Domäne` | Both | Identifies groupings as domains | "Sales Domain", "Finance Domain" |
| `Fachbereich` | applist | Identifies groupings as functional areas | "Customer Management", "Accounting" |

### For Application Components

| Specialization | Used By | Purpose | Example |
|----------------|---------|---------|---------|
| `Geschäftsanwendung` | applist | Business application | "CRM System", "ERP" |
| `Register` | applist | Registry/catalog application | "Customer Register", "Product Catalog" |
| `Querschnittsanwendung` | applist | Cross-cutting application | "Authentication Service", "Logging Service" |

**Important:**
- Specializations are **case-sensitive**
- Only exact matches are recognized
- Application components without recognized specializations are excluded from applist

---

## Properties

Properties are key-value pairs attached to elements or relationships.

### Model-Level Properties

| Property | Element | Used By | Purpose | Example |
|----------|---------|---------|---------|---------|
| `baseline` | Model | comatrix | Name of baseline model for comparison | "MyProject_v1.0" |

**Usage:**
1. Add property "baseline" to your current model
2. Set value to the name of another open model
3. Run comatrix → outputs show changes (green/yellow/red highlighting)

### Relationship Properties

| Property | Relationship | Used By | Purpose | Example |
|----------|-------------|---------|---------|---------|
| `Schnittstelle` | Triggering | comatrix | Interface/API name | "REST API Orders" |

**Important:**
- Can have multiple values with the same key "Schnittstelle"
- Each value creates a separate row in the connectivity matrix
- Empty/missing property shows as "N/A" in output

**Example with multiple interfaces:**
```
Triggering Relationship: "NST_MultiInterface"
  Property "Schnittstelle": "API v1"
  Property "Schnittstelle": "API v2"
  Property "Schnittstelle": "Kafka Topic"
```

**Result:** Creates 3 separate rows in comatrix (one per interface)

---

## Naming Conventions

### Triggering Relationships

**Convention:** All triggering relationships analyzed by comatrix MUST start with `NST_` prefix.

**Examples:**
- ✅ `NST_OrderPlacement`
- ✅ `NST_CustomerQuery`
- ✅ `NST_PaymentProcessing`
- ❌ `OrderPlacement` (no NST_ prefix → ignored)
- ❌ `nst_lowercase` (case-sensitive → ignored)

**Rationale:** Filters out irrelevant triggering relationships from your model.

### Elements and Groupings

**No specific naming convention required** - any name is valid. However, names should be:
- **Unique within type**: Duplicate names may cause confusion in reports
- **Descriptive**: Names appear directly in Excel outputs
- **Consistent**: Use similar patterns for better readability

---

## Model Structure Requirements

### Domain/Fachbereich Assignment

The scripts traverse aggregation/composition relationships to find domains and Fachbereich. Understanding this traversal is crucial for correct outputs.

#### Basic Structure (Single Level)

```
Grouping: "Sales Domain"
  Specialization: "Domäne"
    │
    └── aggregation ──→ Application Component: "Customer Portal"
                         Result: domain = "Sales Domain"
```

#### Nested Structure (Multi-Level)

```
Grouping: "Customer Management"
  Specialization: "Fachbereich"
    │
    └── aggregation ──→ Grouping: "Sales Domain"
                         Specialization: "Domäne"
                           │
                           └── aggregation ──→ Application Component: "CRM"
                                                Result: 
                                                  - domain = "Sales Domain"
                                                  - Fachbereich = "Customer Management"
```

**Key Points:**
- Traversal goes UP from application component through ALL parent groupings
- Collects ALL groupings with target specialization
- Handles nested structures of arbitrary depth

#### Multiple Domains/Fachbereich

```
Grouping: "Sales Domain"          Grouping: "Finance Domain"
  Specialization: "Domäne"         Specialization: "Domäne"
    │                                 │
    └── aggregation ──┐           ┌──┘
                      │           │
                      ▼           ▼
              Application Component: "Order System"
              Result: domain = "Sales Domain, Finance Domain"
```

**Output:** All values are comma-separated in Excel: `"Sales Domain, Finance Domain"`

#### Cycle Detection

```
Grouping A ──aggregation──→ Grouping B
    ▲                           │
    │                           │
    └─────── aggregation ───────┘
```

**Result:** "cycle" is written instead of domain/Fachbereich name

**Why this matters:**
- Prevents infinite loops
- Indicates a modeling error that should be fixed
- Visible in output for troubleshooting

### Connectivity Matrix Structure

```
Application Component A (target of NST_*)
  ├── Triggering Relationship: "NST_Service1"
  │     Property "Schnittstelle": "API X"
  │     Source: Application Component B
  │
  └── Triggering Relationship: "NST_Service2"
        Property "Schnittstelle": "API Y"
        Source: Application Component C

Result in comatrix.xlsx:
  Row 1: A | API X | [intern/extern] | x in column B
  Row 2: A | API Y | [intern/extern] | x in column C
```

---

## Impact on Outputs

### Comatrix.xlsx Output

**What influences the rows?**
- Each combination of (A-element, Schnittstelle) creates one row
- A-elements = targets of NST_* triggering relationships
- Schnittstellen = values of "Schnittstelle" property on relationships

**What influences the columns?**
- B-elements = sources of NST_* triggering relationships
- Column headers show B-element names

**What influences cell content?**
- "x" appears when relationship exists between (A-element + Schnittstelle) and B-element
- Cell color (in comparison mode):
  - Green: New connection (not in baseline)
  - Yellow/Orange: Changed connection
  - Red: Removed connection (in baseline but not current)

**What influences domain columns?**
- Domain column for A-elements shows result of findDomain(target)
- Domain row for B-elements shows result of findDomain(source)
- Domains from grouping elements with specialization "Domäne"

**What influences intern/extern classification?**
- "intern": A-element and B-element have same domain
- "extern": A-element and B-element have different domains
- Empty if domain(s) not found

**What influences comparison colors?** (baseline mode only)
- Model property "baseline" must point to an open model
- Compares relationship existence between baseline and current model
- Element and connection presence/absence determines colors

### Applist.xlsx Output

**What influences which applications appear?**
- Only application components with specialization:
  - "Geschäftsanwendung"
  - "Register"
  - "Querschnittsanwendung"
- All three types are included

**What influences the "Typ" column?**
- Directly copied from element's specialization

**What influences the "Domäne" column?**
- Result of findDomain() function
- All groupings with specialization "Domäne" in the hierarchy
- Comma-separated if multiple
- "cycle" if circular grouping reference
- "(keine Domäne)" if none found

**What influences the "Fachbereich" column?**
- Result of findFachbereich() function
- All groupings with specialization "Fachbereich" in the hierarchy
- Comma-separated if multiple
- "cycle" if circular grouping reference
- "(kein Fachbereich)" if none found

**What influences the sort order?**
1. First by Fachbereich (alphabetically, empty values last)
2. Then by Domäne (alphabetically, empty values last)
3. Then by Typ (alphabetically)
4. Finally by application name (alphabetically)

---

## Checklist: Ensuring Correct Outputs

### For Comatrix

- [ ] Triggering relationships named with `NST_` prefix
- [ ] Property "Schnittstelle" set on triggering relationships
- [ ] Application components connected via triggering relationships
- [ ] Groupings with specialization "Domäne" created
- [ ] Application components grouped (aggregation/composition) into domains
- [ ] (Optional) Model property "baseline" set for comparison mode

### For Applist

- [ ] Application components have specialization: "Geschäftsanwendung", "Register", or "Querschnittsanwendung"
- [ ] Groupings with specialization "Domäne" created
- [ ] Groupings with specialization "Fachbereich" created
- [ ] Application components grouped into domains
- [ ] Application components grouped into Fachbereich
- [ ] No circular grouping references (or accept "cycle" in output)

### Common Issues

| Issue | Symptom | Solution |
|-------|---------|----------|
| Empty comatrix | "⚠ No NST_* triggering relationships found" | Add/rename triggering relationships with NST_ prefix |
| Empty applist | "⚠ No application components with the specified specializations found" | Set specialization on application components |
| Missing domains | "(keine Domäne)" in output | Create grouping with specialization "Domäne" and link via aggregation/composition |
| Missing Fachbereich | "(kein Fachbereich)" in output | Create grouping with specialization "Fachbereich" and link via aggregation/composition |
| "cycle" in output | "cycle" appears in domain/Fachbereich column | Fix circular grouping references in model |
| Wrong intern/extern | Incorrect classification | Check domain assignment and ensure consistent grouping |
| Missing connections | Expected "x" doesn't appear | Verify NST_* relationship exists with correct source/target |

---

## Example Model Structure

### Complete Working Example

```
Model: "Enterprise Architecture"
  Property "baseline": "Enterprise_Architecture_v1"

  Grouping: "Customer Management"
    Specialization: "Fachbereich"
    
    Grouping: "Sales Domain"
      Specialization: "Domäne"
      
      Application Component: "Customer Portal"
        Specialization: "Geschäftsanwendung"
        
      Application Component: "Order System"
        Specialization: "Geschäftsanwendung"

  Grouping: "Finance Domain"
    Specialization: "Domäne"
    
    Application Component: "Payment Gateway"
      Specialization: "Querschnittsanwendung"

  Triggering Relationship: "NST_PlaceOrder"
    Source: "Customer Portal"
    Target: "Order System"
    Property "Schnittstelle": "REST API Orders"
    Property "Schnittstelle": "WebSocket Updates"

  Triggering Relationship: "NST_ProcessPayment"
    Source: "Order System"
    Target: "Payment Gateway"
    Property "Schnittstelle": "SOAP Payment Service"
```

**Expected comatrix.xlsx output:**
| Domäne | Anwendungssystem | Schnittstelle | intern/extern | Customer Portal | Order System | Payment Gateway |
|--------|------------------|---------------|---------------|-----------------|--------------|-----------------|
| | **Legend** | **Legend** | **Legend** | | | |
| Sales Domain | Order System | REST API Orders | intern | x | | |
| Sales Domain | Order System | WebSocket Updates | intern | x | | |
| Finance Domain | Payment Gateway | SOAP Payment Service | extern | | x | |

**Expected applist.xlsx output:**
| Anwendung | Typ | Domäne | Fachbereich |
|-----------|-----|---------|-------------|
| Customer Portal | Geschäftsanwendung | Sales Domain | Customer Management |
| Order System | Geschäftsanwendung | Sales Domain | Customer Management |
| Payment Gateway | Querschnittsanwendung | Finance Domain | (kein Fachbereich) |

---

## Advanced Topics

### Multiple Property Values

The "Schnittstelle" property can have multiple values with the same key. This is an Archi feature where you can add the same property name multiple times with different values.

**How to add:**
1. Select triggering relationship
2. Properties tab → Add property "Schnittstelle" with value "API v1"
3. Add another property "Schnittstelle" with value "API v2"

**Result:** Each value creates a separate row in the connectivity matrix.

### Nested Groupings

You can nest groupings arbitrarily deep. The scripts traverse the entire hierarchy upward from the application component.

**Example:**
```
Fachbereich: "IT Operations"
  └──→ Domäne: "Infrastructure"
         └──→ Domäne: "Cloud Services" (nested domain)
                └──→ Application: "Kubernetes Cluster"
```

**Result:** Application shows BOTH domains comma-separated: "Infrastructure, Cloud Services"

### Baseline Comparison Mode

When the model property "baseline" is set:
1. Script searches for an open model with that name
2. If found, analyzes both models
3. Merges results and highlights differences

**Color coding:**
- Green: Element/connection in current, not in baseline (new)
- Yellow/Orange: Element in both, but connections changed
- Red: Element/connection in baseline, not in current (removed)

**Setup:**
1. Open baseline model in Archi
2. Open current model in Archi
3. Add property "baseline" to current model with baseline model's name
4. Run comatrix script

---

## Summary

The comatrix and applist scripts analyze your Archi model based on:

1. **Element types**: Application components, groupings
2. **Relationship types**: Triggering (NST_*), aggregation, composition
3. **Specializations**: Domäne, Fachbereich, Geschäftsanwendung, Register, Querschnittsanwendung
4. **Properties**: Schnittstelle (on relationships), baseline (on model)
5. **Structure**: Hierarchical grouping via aggregation/composition

Understanding these patterns helps you structure your Archi model to produce accurate, meaningful Excel reports that support enterprise architecture documentation and analysis.
