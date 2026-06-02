# ⚡ Jour 4 / 10 — Formules Dynamiques Excel

> **Série : 10 Days of Excel** · Jour 4/10  
> Concepts : FILTRE · UNIQUE · TRIER · TRANSPOSE · Spill

---

## 📁 Structure du fichier

```
formules_dynamiques_jour4.xlsx
│
├── DONNÉES              ← 30 ventes (vendeur, région, produit, montant, objectif)
├── FORMULES DYNAMIQUES  ← Les 4 formules avec exemples concrets
└── EXERCICES            ← 5 exercices progressifs avec solutions
```

---

## ⚡ C'est quoi une formule dynamique ?

**Avant (Excel 2019 et antérieur) :**  
Tu écrivais `=B2` en C2, tu tirais vers le bas jusqu'à C100. 100 cellules à gérer.

**Après (Excel 2021 / Microsoft 365) :**  
Tu écris `=FILTRE(B2:B100, ...)` en C2 — Excel déverse automatiquement tous les résultats. On appelle ça le **spill** (déversement).

Le `#` dans une référence comme `C2#` signifie *"toute la plage déversée depuis C2"*.

---

## 🔑 Les 4 formules

### 1. FILTRE
```excel
=FILTRE(tableau, condition, [si_vide])

# Exemple — ventes de la région Île-de-France
=FILTRE(DONNÉES!B2:H31, DONNÉES!C2:C31="Île-de-France", "Aucun résultat")

# Multi-critères ET (les deux conditions vraies)
=FILTRE(tableau, (col1="Mobile") * (col2>3000))

# Multi-critères OU (au moins une condition vraie)
=FILTRE(tableau, (col1="Mobile") + (col1="Audio"))
```

### 2. UNIQUE
```excel
=UNIQUE(tableau, [par_colonne], [seulement_une_fois])

# Liste vendeurs sans doublons
=UNIQUE(DONNÉES!B2:B31)

# Combiné avec TRIER
=TRIER(UNIQUE(DONNÉES!D2:D31))
```

### 3. TRIER
```excel
=TRIER(tableau, [colonne], [ordre], [par_colonne])

# Trier par montant décroissant (col 5, ordre -1)
=TRIER(DONNÉES!B2:F31, 5, -1)

# Trier par 2 colonnes : montant desc puis vendeur asc
=TRIER(DONNÉES!B2:F31, {5,1}, {-1,1})
```
- Ordre `1` = croissant · Ordre `-1` = décroissant

### 4. TRANSPOSE
```excel
=TRANSPOSE(tableau)

# Colonne → Ligne
=TRANSPOSE(B2:B6)

# Combiner avec UNIQUE
=TRANSPOSE(UNIQUE(DONNÉES!B2:B31))
```

---

## 📊 Combinaisons puissantes

| Formule | Résultat |
|---------|---------|
| `=TRIER(UNIQUE(D2:D31))` | Produits uniques triés A→Z |
| `=TRIER(FILTRE(B2:F31, E2:E31="Mobile"), 5, -1)` | Ventes Mobile triées par montant |
| `=TRANSPOSE(UNIQUE(B2:B31))` | Vendeurs uniques en ligne |

---

## 💡 Règles à retenir

- `*` entre deux conditions = **ET** logique
- `+` entre deux conditions = **OU** logique
- `tableau#` = référence à toute la plage déversée
- Ces formules nécessitent **Excel 2021 ou Microsoft 365**
