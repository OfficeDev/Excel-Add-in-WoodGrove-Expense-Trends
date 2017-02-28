# <a name="excel-add-in-js-woodgrove-expense-trends"></a>Excel-Add-in-JS-WoodGrove-Expense-Trends

Le complément de tendances de dépenses bancaires WoodGrove montre comment vous pouvez utiliser la nouvelle API JavaScript pour Microsoft Excel 2016 en vue de créer un complément Excel attrayant. Avec ce complément, vous pouvez importer des transactions de dépenses dans le classeur, créer des tableaux de bord et des suivis, afficher et analyser les tendances, et assurer le suivi des transactions spéciales telles que les dons. L’exemple fournit deux expériences : une avec le volet Office et l’autre avec les commandes de complément. Les figures suivantes présentent les écrans principaux de ce complément.

![Complément de tendances de dépenses bancaires WoodGrove - Ruban] (images/woodgrove_taskpane_ribbon.PNG)

![Complément de tendances de dépenses bancaires WoodGrove - Volet Office initial] (images/woodgrove_taskpane_import.PNG)

![Complément de tendances de dépenses bancaires WoodGrove - Feuille de transactions] (images/woodgrove_taskpane_data.PNG)

![Complément de tendances de dépenses bancaires WoodGrove - Tableau de bord] (images/woodgrove_taskpane_dashboard.PNG)

![Complément de tendances de dépenses bancaires WoodGrove - Suivi des dons] (images/woodgrove_taskpane_donations.PNG)

## <a name="table-of-contents"></a>Sommaire

* [Conditions préalables](#prerequisites)
* [Exécution du projet](#run-the-project)
* [Ressources supplémentaires](#additional-resources)

## <a name="prerequisites"></a>Conditions préalables

Vous devez disposer des éléments suivants :

* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Outils de développement Office pour Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* Excel 2016, version 6769.2011 ou ultérieure

## <a name="run-the-project"></a>Exécuter le projet

1. Copiez le projet dans un dossier local. Vérifiez que le chemin d’accès de fichier n’est pas trop long. Si c’est le cas, il est possible que vous rencontriez une erreur dans Visual Studio lorsque ce dernier tente d’installer les packages NuGet nécessaires pour le projet. 
2. Ouvrez le fichier `WoodGrove Expense Trends.sln` dans Visual Studio. 
3. Appuyez sur F5 pour créer et déployer l’exemple de complément. Excel démarre et, selon la version d’Excel 2016 dont vous disposez, le complément charge un onglet personnalisé intitulé WoodGrove dans le ruban ou s’ouvre dans un volet Office à droite de la feuille de calcul, comme illustré dans les figures suivantes.

![Complément de tendances de dépenses bancaires WoodGrove - Volet Office initial] (images/woodgrove_taskpane_ribbon.PNG)

![Complément de tendances de dépenses bancaires WoodGrove - Volet Office initial] (images/woodgrove_taskpane_import.PNG)

## <a name="additional-resources"></a>Ressources supplémentaires

* [Centre de développement Office](http://dev.office.com/)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft. Tous droits réservés.

