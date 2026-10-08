# Documentation française

[English documentation](../en/README.md)

## Commencer ici

- [Carte de lecture de l'architecture](ARCHITECTURE_READING_GUIDE.md) : comprendre les domaines, workflows, propriétaires et garde-fous.
- [Guide de maintenance](MAINTENANCE_GUIDE.md) : modifier le projet sans casser ses contrats.
- [Glossaire du projet](PROJECT_GLOSSARY.md) : vocabulaire officiel utilisé dans le code et la documentation.

## Utiliser ProjectEngine v1.3.0

- **Premier démarrage :** l'accueil propose **Start New Project** ou **Import Previous Project** si le classeur est vide et modifiable. Le choix est conservé par le registre d'acquittement existant ; Full Reset peut réafficher l'accueil.
- **Ruban Excel :** l'onglet **ProjectEngine**, embarqué dans le classeur, rassemble mises à jour, commandes Gantt, paramètres, navigation, réinitialisation et import. Les commandes sont contextuelles ; aucune extension externe n'est requise.
- **Importer un ancien projet :** sélectionner **Import Previous Project** dans l'accueil ou le Ribbon. Les données prises en charge remplacent celles de la destination ; la source reste intacte. Vérifier ensuite WBS/CONSTRAINTS, puis lancer soi-même Planning Update ou Full Update. **Effectuer soi-même une copie de sécurité du classeur de destination avant l'import.**
- **Navigation Gantt :** **Go to Selected Task** et **Show Full Timeline** exploitent l'owner Gantt existant. Les très grands plannings peuvent dépasser les limites physiques de zoom d'Excel.

La [carte d'architecture](ARCHITECTURE_READING_GUIDE.md) décrit les responsabilités et parcours ; le [guide de maintenance](MAINTENANCE_GUIDE.md) précise les précautions pour les évolutions.

## Parcours recommandé

1. Lire la carte de lecture pour obtenir une vue générale en 15 minutes.
2. Consulter le glossaire avant de créer ou renommer un composant.
3. Utiliser le guide de maintenance pour choisir la frontière et les validations adaptées.
4. Consulter les en-têtes FR/EN des modules et procédures pour les contrats détaillés du code actuel.
