# <a name="excel-add-in-js-woodgrove-expense-trends"></a>Excel-Add-in-JS-WoodGrove-Expense-Trends

O suplemento de Tendências de Despesas do Banco WoodGrove demonstra como você pode usar a nova API JavaScript do Microsoft Excel 2016 para criar um atraente suplemento do Excel. Com as Tendências de Despesas, você pode importar transações de despesas para a pasta de trabalho, criar painéis e rastreadores, exibir e analisar tendências e controlar transações especiais como doações para caridade e itens de acompanhamento. O exemplo oferece duas experiências: uma com o painel de tarefas e outra com os comandos do suplemento. As figuras a seguir mostram as telas principais desse suplemento.

![Suplemento de Tendências de Despesas do Banco WoodGrove - Faixa de Opções] (images/woodgrove_taskpane_ribbon.PNG)

![Suplemento de Tendências de Despesas do Banco WoodGrove - Painel de tarefas inicial] (images/woodgrove_taskpane_import.PNG)

![Suplemento de Tendências de Despesas do Banco WoodGrove - Planilha de transações] (images/woodgrove_taskpane_data.PNG)

![Suplemento de Tendências de Despesas do Banco WoodGrove - Painel] (images/woodgrove_taskpane_dashboard.PNG)

![Suplemento de Tendências de Despesas do Banco WoodGrove - Rastreador de Doações] (images/woodgrove_taskpane_donations.PNG)

## <a name="table-of-contents"></a>Sumário

* [Pré-requisitos](#prerequisites)
* [Executar o projeto](#run-the-project)
* [Recursos adicionais](#additional-resources)

## <a name="prerequisites"></a>Pré-requisitos

Você precisará do seguinte:

* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Office Developer Tools para Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* Excel 2016, versão 6769.2011 ou posterior

## <a name="run-the-project"></a>Executar o projeto

1. Copie o projeto para uma pasta local. Certifique-se de que o caminho do arquivo não seja muito longo, caso contrário, você pode encontrar um erro no Visual Studio ao tentar instalar os pacotes do NuGet necessários para o projeto. 
2. Em seguida, abra o `WoodGrove Expense Trends.sln` no Visual Studio. 
3. Pressione F5 para criar e implantar o suplemento de exemplo. O Excel é iniciado e, dependendo da versão do Excel 2016 que você tiver, o suplemento carregará uma guia personalizada chamada WoodGrove na faixa de opções ou abrirá um painel de tarefas à direita da planilha, conforme mostrado nas figuras a seguir.

![Suplemento de Tendências de Despesas do Banco WoodGrove - Painel de tarefas inicial] (images/woodgrove_taskpane_ribbon.PNG)

![Suplemento de Tendências de Despesas do Banco WoodGrove - Painel de tarefas inicial] (images/woodgrove_taskpane_import.PNG)

## <a name="additional-resources"></a>Recursos adicionais

* [Centro de Desenvolvimento do Office](http://dev.office.com/)

## <a name="copyright"></a>Copyright
Copyright © 2016 Microsoft. Todos os direitos reservados.

