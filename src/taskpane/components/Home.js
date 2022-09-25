import React, { Component } from 'react';
import './home.css'


class Home extends Component {
    constructor(props){
        super(props)
        this.state = {
            
        }

       }
      async   createTable() {
         Excel.run(async (context) => {
    
            // TABLEAUX.
            const currentWorksheet = context.workbook.worksheets.add("Bilan sociale");
            currentWorksheet.activate();
            currentWorksheet.name="Bilan social";
            const expensesTable = currentWorksheet.tables.add("A1:F1", true /*hasHeaders*/);
            expensesTable.name = "ExpensesTable";

            // BILAN ENV
            const newSheet = context.workbook.worksheets.add("Bilan environnemental");
            newSheet.activate();
            const bilanEnv = newSheet.tables.add("A1:E1", true /*hasHeaders*/);
            bilanEnv.name = "BilanEnv";

             
             bilanEnv.getHeaderRowRange().values =
             [["Type d'indicateurs", "Input à saisir","Valeur", "Champ calculé", "Valeur_de_dette"]];
        
        bilanEnv.rows.add(null, [
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets inertes produite (gravat, sable, tuiles, béton, ciment, carrelage..)", "1", ""," "],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets inertes recyclée", "1", "Taux de déchets inertes non recyclés","=((C2-C3)/C2)"],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets industriels non dangereux produite", "1", "",""],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets industriels non dangereux recyclée", "1", "Taux de déchets industriels non dangereux non recyclés","=((C4-C5)/C4)"],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets industriels dangereux produite", "1", "Limite moyenne à respecter","1.5%"],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets plastiques produite", "1", "",""],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets plastiques recyclée", "1", "Taux de déchets plastiques non recyclés","=((C6-C7)/C6)"],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets fermentescibles produite", "1", "",""],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets fermentescibles recyclée", "1", "Taux de déchets fermentescibles non recyclés","=((C8-C9)/C8)"],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets d’équipements electroniques et éléctriques produite", "1", "",""],
            ["Élimination des déchets et efforts de l’entreprise pour en limiter la quantité", "Quantité de déchets d’équipements electroniques et éléctriques recyclée", "1", "Taux de déchets d’équipements electroniques et electriques non recyclés","=((C10-C11)/C10)"],
            

            ["Lutte contre la pollution du sol et des eaux","Concentration de produits chimiques déversée dans le m² de sol",  "1", "Limite moyenne à respecter",""],
            ["Lutte contre la pollution du sol et des eaux","Concentration de produits chimiques déversée dans le litre d’eau (lac, mer, étang..etc)",  "1", "Limite moyenne à respecter",""],
            ["Lutte contre la pollution du sol et des eaux","Concentration de déchets solides déversée dans le m² de sol",  "1", "Limite moyenne à respecter",""],
            ["Lutte contre la pollution du sol et des eaux","Concentration de déchets solides déversée dans le litre d’eau (lac, mer, étang..)",  "1", "Limite moyenne à respecter",""],
            ["Lutte contre la pollution du sol et des eaux","Fréquence d’exploitation agricole du m² de sol",  "1", "Limite moyenne à respecter","75%"],
            ["Lutte contre la pollution du sol et des eaux","Quantité d’eau utilisée pour arrosage agricole par heure et par m²",  "1", "Limite moyenne à respecter","900000"],
            
            ["Préservation de la qualité de l’air et du climat","Quantité de CO2 générée par les processus internes de fabrication",  "1", "Limite moyenne à respecter","1000 mg/kg"],
            ["Préservation de la qualité de l’air et du climat","Quantité de CO2 générée par les déplacements sur chantier",  "1", "Limite moyenne à respecter","5 g/km"],
            
            ["Réduction des nuisances sonores","Nombre de plaintes reçues des habitants en périphérie pour cause de nuisance sonore",  "1", "",""],
            ["Réduction des nuisances sonores","Nombre de certificats maladies de collaborateurs pour cause d’impact de nuisance sonore",  "1", "",""],
            ["Réduction des nuisances sonores","Nombre de procès juridique pour cause de nuisance sonore",  "1", "",""],
            ["Réduction des nuisances sonores","Nombre de mesures décidées pour atténuer les nuisances sonores",  "1", "",""],
            ["Réduction des nuisances sonores","Nombre de mesures appliquées pour atténuer les nuisances sonores",  "1", "Taux de nuisance sonore à résoudre","=(1-((C24+C25)/(C21+C22+C23)))"],

            ["Protection de la biodiversité et du paysage","Nombre d’arbres déracinés",  "1", "",""],
            ["Protection de la biodiversité et du paysage","Nombre d’arbres plantés",  "1", "",""],
            ["Protection de la biodiversité et du paysage","Nombre de nouveaux bâtiments construits",  "1", "",""],
            ["Protection de la biodiversité et du paysage","Nombre de bâtiments détruits",  "1", "Quantité restante d’arbres à planter","=(1-(C27)/(C26+C28+C29))"],
            ["Protection de la biodiversité et du paysage","Quantité de déchets de chantier de construction produite",  "1", "",""],
            ["Protection de la biodiversité et du paysage","Quantité de déchets de chantier de construction recyclés",  "1", "Quantité de déchets de chantier de construction non recyclée","=((C30-C31)/C30)"],
        ]);
            
     
            // BILAN SOCIAL
            expensesTable.getHeaderRowRange().values =
            [["Mois", "Type d'indicateurs", "Input à saisir", "Valeur","Champ calculé","Valeur"]];
        
        expensesTable.rows.add(null /*add at the end*/, [
            // JANVIER
            ["Janvier", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D2/D3"],
            ["Janvier", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
            ["Janvier", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D4/D3"],
            ["Janvier", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
            ["Janvier", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D5/D6"],
            ["Janvier", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D2+D4)/D7"],
            ["Janvier", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D9/D8"],
            ["Janvier", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
            ["Janvier", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D10/D9"],
            ["Janvier", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
            ["Janvier", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D11/D12"],
            ["Janvier", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
            ["Janvier", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D13/D14"],
            ["Janvier", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
            ["Janvier", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D15/(251*D16)"],
            ["Janvier", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
            ["Janvier", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D17/D18"],
            ["Janvier", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D19*10^6)/D20"],
            ["Janvier", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
            ["Janvier", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D21*10^3)/D20"],
            ["Janvier", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
            ["Janvier", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D22/D23"],
            ["Janvier", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
            ["Janvier", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D24/D25"],
            ["Janvier", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
            ["Janvier", "Les indicateurs démographiques","Effectif total", "45","Age","=D26/D27"],
            ["Janvier", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D28/D27"],
            ["Janvier", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
            ["Janvier", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D30/D29"],
            ["Janvier", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D30/D31"],
            ["Janvier", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
            ["Janvier", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D32/D33"],
            // FEVRIER
            ["Fevrier", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
            ["Fevrier", "Indicateurs de mobilité", "Nombre d'agents", "200","Taux de sortie","=D34/D35"],
            ["Fevrier", "Indicateurs de mobilité", "Nombre d'agents", "200","Taux d'entrée","=D37/D35"],
            ["Fevrier", "Indicateurs de mobilité", "Nombre d'entrées", "60","Taux d'entrée","=D37/D35"],
            ["Fevrier", "Indicateurs de mobilité", "Nombre d'arrivées", "19","Ratio de remplacement","=D38/D39"],
            ["Fevrier", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D38/D39"],
            ["Fevrier", "Indicateurs de mobilité", "Effectif", "70"," Turn Over","=(D34+D36)/D40"],


            ["Fevrier", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D42/D41"],
            ["Fevrier", "Indicateurs d’intégration", "Nombre de départs annuel", "12","Taux des départs","=D42/D41"],
            ["Fevrier", "Indicateurs d’intégration", "Nombre de départs annuel", "12","Taux des départs volontaires","=D44/D42"],
            ["Fevrier", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D44/D42"],
            ["Fevrier", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","Taux des prorogations de stage","=D45/D46"],
            ["Fevrier", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D45/D46"],


            ["Fevrier", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","Taux d'ancienneté dans l'organisation","=D47/D48"],
            ["Fevrier", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D47/D48"],


            ["Fevrier", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","Taux d'absentéisme maladie","=D49/(251*D50)"],
            ["Fevrier", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D49/(251*D50)"],
            ["Fevrier", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","Taux d'absentéisme maladie de courte durée","=D51/D52"],
            ["Fevrier", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D51/D52"],


            ["Fevrier", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D53*10^6)/D54"],
            ["Fevrier", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","Taux de fréquence des accidents du travail","=(D53*10^6)/D54"],
            ["Fevrier", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","Taux de gravité des accidents de travail","=(D56*10^3)/D54"],
            ["Fevrier", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D56*10^3)/D54"],


            ["Fevrier", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","Taux de départ en formation par catégorie","=D57/D58"],
            ["Fevrier", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D57/D58"],
            ["Fevrier", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","Taux de participation à la formation","=D59/D60"],
            ["Fevrier", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D59/D60"],


            ["Fevrier", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","Age","=D61/D62"],
            ["Fevrier", "Les indicateurs démographiques","Effectif total", "45","Age","=D61/D62"],
            ["Fevrier", "Les indicateurs démographiques","Effectif total", "45","Ancienneté","=D64/D62"],
            ["Fevrier", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D64/D62"],


            ["Fevrier", "Les indicateurs liés aux rémunérations","Effectif", "1","Masse salariale 1","=D66/D65"],
            ["Fevrier", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D66/D65"],
            ["Fevrier", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D66/D67"],
            ["Fevrier", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","Disparité des salaires entre différentes catégories","=D68/D69"],
            ["Fevrier", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D68/D69"],

            // MARS
            ["Mars", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D66/D67"],
            ["Mars", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
            ["Mars", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D68/D67"],
            ["Mars", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
            ["Mars", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D69/D70"],
            ["Mars", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D66+D68)/D71"],
            ["Mars", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D73/D72"],
            ["Mars", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
            ["Mars", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D74/D73"],
            ["Mars", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
            ["Mars", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D75/D76"],
            ["Mars", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
            ["Mars", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D77/D78"],
            ["Mars", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
            ["Mars", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D79/(251*D80)"],
            ["Mars", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
            ["Mars", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D81/D82"],
            ["Mars", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D83*10^6)/D84"],
            ["Mars", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
            ["Mars", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D85*10^3)/D84"],
            ["Mars", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
            ["Mars", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D86/D87"],
            ["Mars", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
            ["Mars", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D88/D89"],
            ["Mars", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
            ["Mars", "Les indicateurs démographiques","Effectif total", "45","Age","=D90/D91"],
            ["Mars", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D92/D91"],
            ["Mars", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
            ["Mars", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D94/D93"],
            ["Mars", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D94/D95"],
            ["Mars", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
            ["Mars", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D96/D97"],
                   
            // AVRIL
            ["Avril", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
            ["Avril", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
            ["Avril", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
            ["Avril", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
            ["Avril", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
            ["Avril", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
            ["Avril", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
            ["Avril", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
            ["Avril", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
            ["Avril", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
            ["Avril", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
            ["Avril", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
            ["Avril", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
            ["Avril", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
            ["Avril", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
            ["Avril", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
            ["Avril", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
            ["Avril", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
            ["Avril", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
            ["Avril", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
            ["Avril", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
            ["Avril", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
            ["Avril", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
            ["Avril", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
            ["Avril", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
            ["Avril", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
            ["Avril", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
            ["Avril", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
            ["Avril", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
            ["Avril", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
            ["Avril", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
            ["Avril", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

            // MAI
            ["Mai", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Mai", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Mai", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Mai", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Mai", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Mai", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Mai", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Mai", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Mai", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Mai", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Mai", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Mai", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Mai", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Mai", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Mai", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Mai", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Mai", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Mai", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Mai", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Mai", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Mai", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Mai", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Mai", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Mai", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Mai", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Mai", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Mai", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Mai", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Mai", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Mai", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Mai", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Mai", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

// JUIN
["Juin", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Juin", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Juin", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Juin", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Juin", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Juin", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Juin", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Juin", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Juin", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Juin", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Juin", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Juin", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Juin", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Juin", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Juin", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Juin", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Juin", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Juin", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Juin", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Juin", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Juin", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Juin", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Juin", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Juin", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Juin", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Juin", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Juin", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Juin", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Juin", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Juin", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Juin", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Juin", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],
    
// JUILLET
["Juillet", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Juillet", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Juillet", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Juillet", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Juillet", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Juillet", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Juillet", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Juillet", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Juillet", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Juillet", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Juillet", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Juillet", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Juillet", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Juillet", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Juillet", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Juillet", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Juillet", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Juillet", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Juillet", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Juillet", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Juillet", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Juillet", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Juillet", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Juillet", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Juillet", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Juillet", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Juillet", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Juillet", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Juillet", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Juillet", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Juillet", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Juillet", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

//AOUT
["Aout", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Aout", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Aout", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Aout", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Aout", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Aout", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Aout", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Aout", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Aout", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Aout", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Aout", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Aout", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Aout", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Aout", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Aout", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Aout", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Aout", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Aout", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Aout", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Aout", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Aout", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Aout", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Aout", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Aout", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Aout", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Aout", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Aout", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Aout", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Aout", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Aout", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Aout", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Aout", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

// SEPTEMBRE
["Septembre", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Septembre", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Septembre", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Septembre", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Septembre", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Septembre", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Septembre", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Septembre", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Septembre", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Septembre", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Septembre", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Septembre", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Septembre", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Septembre", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Septembre", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Septembre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Septembre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Septembre", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Septembre", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Septembre", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Septembre", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Septembre", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Septembre", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Septembre", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Septembre", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Septembre", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Septembre", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Septembre", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Septembre", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Septembre", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Septembre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Septembre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

// OCTOBRE
["Octobre", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Octobre", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Octobre", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Octobre", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Octobre", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Octobre", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Octobre", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Octobre", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Octobre", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Octobre", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Octobre", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Octobre", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Octobre", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Octobre", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Octobre", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Octobre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Octobre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Octobre", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Octobre", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Octobre", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Octobre", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Octobre", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Octobre", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Octobre", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Octobre", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Octobre", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Octobre", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Octobre", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Octobre", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Octobre", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Octobre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Octobre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

// NOVEMBRE
["Novembre", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Novembre", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Novembre", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Novembre", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Novembre", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Novembre", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Novembre", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Novembre", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Novembre", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Novembre", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Novembre", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Novembre", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Novembre", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Novembre", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Novembre", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Novembre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Novembre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Novembre", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Novembre", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Novembre", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Novembre", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Novembre", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Novembre", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Novembre", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Novembre", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Novembre", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Novembre", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Novembre", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Novembre", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Novembre", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Novembre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Novembre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],

// DECEMBRE
["Decembre", "Indicateurs de mobilité", "Nombre de sorties", "19","Taux de sortie/d'entrée","=D34/D35"],
["Decembre", "Indicateurs de mobilité", "Nombre d'agents", "200","",""],
["Decembre", "Indicateurs de mobilité", "Nombre d’entrées", "60","Taux de sortie/d'entrée","=D36/D35"],
["Decembre", "Indicateurs de mobilité", "Nombre d'arrivées", "19","",""],
["Decembre", "Indicateurs de mobilité", "Nombre de départs", "3","Ratio de remplacement","=D37/D38"],
["Decembre", "Indicateurs de mobilité", "Effectif", "70","Turn Over","=(D34+D36)/D39"],
["Decembre", "Indicateurs d’intégration", "Nombre moyen d’agents", "60","Taux des départs","=D41/D40"],
["Decembre", "Indicateurs d’intégration", "Nombre de départs annuel", "12","",""],
["Decembre", "Indicateurs d’intégration", "Nombre de démissions / détachement sur l’année", "6","Taux des départs volontaires","=D42/D41"],
["Decembre", "Indicateurs d’intégration", "Nombre des prorogations de stage", "10","",""],
["Decembre", "Indicateurs d’intégration", "Total des mises en stage", "7","Taux des prorogations de stage","=D43/D44"],
["Decembre", "Les indicateurs liés à l’emploi", "Nombre d'agents ayant moins de x ans dans la collectivité ", "20","",""],
["Decembre", "Les indicateurs liés à l’emploi", "Effectif moyen", "100","Taux d'ancienneté dans l'organisation","=D45/D46"],
["Decembre", "Les indicateurs liés au risque maladie", "Nombre de jours d'absence en jour ouvrée ", "70","",""],
["Decembre", "Les indicateurs liés au risque maladie", "Nomdre d'agents de l'effectif", "150","Taux d'absentéisme maladie","=D47/(251*D48)"],
["Decembre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie moins de 4 jours", "50","",""],
["Decembre", "Les indicateurs liés au risque maladie", "Nombre d'arrêt maladie", "60","Taux d'absentéisme maladie de courte durée","=D49/D50"],
["Decembre", "Les indicateurs liés au risque professionnel", "Nombre d'accidents de travail avec arrêt ", "20","Taux de fréquence des accidents du travail","=(D51*10^6)/D52"],
["Decembre", "Les indicateurs liés au risque professionnel", "Nombre d'heures travaillées", "8000","",""],
["Decembre", "Les indicateurs liés au risque professionnel", "Nombre de jours d'arrêt de travail ", "60","Taux de gravité des accidents de travail","=(D53*10^3)/D52"],
["Decembre", "Les indicateurs de la formation professionnelle","Nombre d'agents par catégorie partis en formation en cours d'année", "7","",""],
["Decembre", "Les indicateurs de la formation professionnelle", "Effectif de la catégorie hiérarchique","67","Taux de départ en formation par catégorie","=D54/D55"],
["Decembre", "Les indicateurs de la formation professionnelle", "Montant des dépenses consacrées a la formation","10000","",""],
["Decembre", "Les indicateurs de la formation professionnelle", "Masse salariale","900000","Taux de participation à la formation","=D56/D57"],
["Decembre", "Les indicateurs démographiques","Agents de plus de 55 ans", "12","",""],
["Decembre", "Les indicateurs démographiques","Effectif total", "45","Age","=D58/D59"],
["Decembre", "Les indicateurs démographiques", "Somme des anciennetés des agents","14","Ancienneté","=D60/D59"],
["Decembre", "Les indicateurs liés aux rémunérations","Effectif", "1","",""],
["Decembre", "Les indicateurs liés aux rémunérations","Frais de personnel", "2","Masse salariale 1","=D62/D61"],
["Decembre", "Les indicateurs liés aux rémunérations","Budget de fonctionnement", "3","Masse salariale 2","=D62/D63"],
["Decembre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus élevés", "4","",""],
["Decembre", "Les indicateurs liés aux rémunérations","Salaire des 10 % les plus bas", "5","Disparité des salaires entre différentes catégories","=D64/D65"],



]);
            // Ajuster le tableau et les formules.
            
            expensesTable.getRange().format.autofitColumns();
            expensesTable.getRange().format.autofitRows();
            bilanEnv.getRange().format.autofitColumns();
            bilanEnv.getRange().format.autofitRows();
            expensesTable.columns.getItemAt(5).getRange().numberFormat = [['%#,##0.00']];
            bilanEnv.columns.getItemAt(4).getRange().numberFormat = [['%#,##0.00']];
            bilanEnv.columns.getItemAt(4).getRange("E25:E31").numberFormat = [['%#,##0.00']];


    
            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
  
    
    render() {
        return (
            <div className='home'>
                <div className='para'>
            <p>En cliquant sur télécharger, deux nouvelles pages s'ajouteront à votre fichier Excel actuel.
Vous pourrez vous baser sur ce template pour intégrer la comptabilité sociale et
 environnementale dans les calculs comptables de votre organisme.  </p>
        </div>
        <div>
            {/* <img src={rahh} alt="template image"/> */}
            <button className="but" id="create-table" onClick={this.createTable}>Télécharger le template</button>
            
        </div>
                
            </div>
        );
    }
}

export default Home;