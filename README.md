# Suivi d'heures de travail — Excel + VBA

Je travaille en restauration et je voulais un moyen simple de suivre mes heures,
calculer ma paie estimée et savoir si je m'approche des 40h dans la semaine.

Au début je notais tout dans un carnet... pas idéal. J'ai décidé de faire
quelque chose en VBA puisque je connaissais déjà les bases.

Ce projet est un peu plus avancé que mon premier (suivi de budget) —
j'ai appris à manipuler des heures et des dates en VBA, ce qui n'est
pas aussi simple que je pensais au départ.

## Ce que ça fait

- Enregistrer un quart de travail (date, heure début, heure fin)
- Calculer automatiquement les heures travaillées
- Estimer la paie brute selon un taux horaire configurable
- Afficher un résumé des heures par semaine
- Avertir si on dépasse 40h dans la semaine

## Structure de la feuille

| Colonne | Contenu |
|---|---|
| Date | Date du quart |
| Début | Heure de début (ex: 16:00) |
| Fin | Heure de fin (ex: 22:30) |
| Heures | Calculé automatiquement |
| Paie estimée | Heures × taux horaire |
| Note | Remarques (ex: "férié", "remplacement") |

## Ce que j'ai appris en faisant ce projet

- Utiliser `Const` pour les valeurs qui ne changent pas (taux horaire)
- Que les heures en VBA sont des fractions de journée — pas des nombres entiers
- Le bug des quarts qui finissent après minuit (ex: 22h00 → 00h30) :
  si on fait juste Fin - Début on obtient un nombre négatif
- `DateDiff` pour calculer des écarts entre deux dates
- `Weekday()` et `DateAdd()` pour trouver le début d'une semaine

---
*Projet personnel — juin 2025*
