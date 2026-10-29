# Documentation française

[English documentation](../en/README.md)

## Commencer ici

- [Carte de lecture de l'architecture](ARCHITECTURE_READING_GUIDE.md) : comprendre les domaines, workflows, propriétaires et garde-fous.
- [Guide de maintenance](MAINTENANCE_GUIDE.md) : modifier le projet sans casser ses contrats.
- [Glossaire du projet](PROJECT_GLOSSARY.md) : vocabulaire officiel utilisé dans le code et la documentation.

## Utiliser ProjectEngine v1.5.0

- **Premier démarrage :** l'accueil propose **Start New Project**, **Import Previous Project** ou **Discover ProjectEngine** si le classeur est vide et modifiable. Le choix est conservé par le registre d'acquittement existant ; Full Reset peut réafficher l'accueil.
- **Ruban Excel :** l'onglet **ProjectEngine**, embarqué dans le classeur, rassemble mises à jour, commandes Gantt, paramètres, navigation, réinitialisation et import. Les commandes sont contextuelles ; aucune extension externe n'est requise.
- **Importer un ancien projet :** sélectionner **Import Previous Project** dans l'accueil ou le Ribbon. Les données prises en charge remplacent celles de la destination ; la source reste intacte. Vérifier ensuite WBS/CONSTRAINTS, puis lancer soi-même Planning Update ou Full Update. **Effectuer soi-même une copie de sécurité du classeur de destination avant l'import.**
- **Tutoriel interactif :** ouvrir **Tutoriel** dans le Ribbon ou **Discover ProjectEngine** depuis l’accueil. **Découverte** présente les fonctions sans opération automatique ; **Pratique** permet les vraies saisies et les commandes normales avec confirmation. Les deux modes disposent de repères de reprise indépendants sur 29 étapes. Aucune donnée de démonstration n’est injectée.
- **Fenêtres modernisées :** accueil, messages et tutoriel bénéficient de nouveaux contrôles et d’une présentation plus lisible. Les messages sont dimensionnés avant ouverture ; texte sélectionnable et défilement conservés, avec repli sur les contrôles natifs si nécessaire.
- **Navigation Gantt :** **Go to Selected Task** et **Show Full Timeline** exploitent l'owner Gantt existant. Les très grands plannings peuvent dépasser les limites physiques de zoom d'Excel.

La [carte d'architecture](ARCHITECTURE_READING_GUIDE.md) décrit les responsabilités et parcours ; le [guide de maintenance](MAINTENANCE_GUIDE.md) précise les précautions pour les évolutions.

## Finance et pilotage multiprojet (nouveau en v1.5.0)

- **FINANCE :** rattacher les activités WBS, saisir Budget et ETC, consulter Réalisé, EAC et Écart ; choisir les profils de distribution **START**, **END** ou **LINEAR**.
- **Journal :** enregistrer les dépenses datées et les avoirs négatifs ; préparer plusieurs modifications dans un tableau Excel natif, avec filtres/tri/collage, puis confirmer ou annuler.
- **Multiprojet :** gérer des projets financiers distincts et des données indépendantes, copiées ou partagées. La vue **Global** consolide sans recompter les enregistrements financiers partagés.
- **FINANCE SUMMARY :** consulter les dépenses prévues/réalisées mensuelles et cumulées. Les opérations futures restent au journal, mais ne sont pas représentées comme réalisé à la date du dernier Refresh Finance.
- **Actualisation explicite :** utiliser **Refresh Finance** ; l'ouverture, le changement de langue et les commandes de planning ne recalculent pas implicitement Finance.
- **Import :** la migration vers **SchemaVersion 3** préserve le domaine financier des fichiers compatibles, ainsi que les formats de saisie WBS et largeurs de colonnes prises en charge. Source non modifiée ; sauvegarder la destination avant import.
- **Réinitialisation complète :** Full Reset efface les données du projet précédent, y compris Finance, extractions, historiques et caches, tout en conservant l'infrastructure et les préférences autorisées.

Il s'agit de **contrôle des coûts de projet**, pas d'une comptabilité générale ni d'un module de facturation. Les profils START/END/LINEAR concernent Finance, pas la S-Curve d'avancement. La vue Global Finance ne fusionne pas plusieurs classeurs de planning.


## Parcours recommandé

1. Lire la carte de lecture pour obtenir une vue générale en 15 minutes.
2. Consulter le glossaire avant de créer ou renommer un composant.
3. Utiliser le guide de maintenance pour choisir la frontière et les validations adaptées.
4. Consulter les en-têtes FR/EN des modules et procédures pour les contrats détaillés du code actuel.
