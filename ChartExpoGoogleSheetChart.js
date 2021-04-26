//<script>
var multiSliderMaxValue = 0;
var startSliderRange = 0;
var endSliderRange = 0;
var isMySavedChartExists = false;
//var chartOpenInEditMode = false;
var selectedChartA1NotationColumns = [];
var exploreSampleDataLinkClicked = false;
var isSampleDataClicked = false;
var isEditModeClicked = false;
var enableCurrentActiveSheetSelection = false;// for enable/disable selection of current active sheet in dropdown
var isEditModeLoadingFromMyChartClick = false;
var inDataSourceScreen = false;
var oldSelectedCellsA1Notations = [];
var headerRowAnnotation = [];
var dataRowsAnnotation = [];
var actualSheetData = [];
var fileName = "ChartExpo"; // these 4 variables state set from editChart and after synching in AutoSync.js as well
var sheetName = "Sheet238";
var sheetNames = [];
var fileId = '';
var sheetId = '';
var selectedScreen = "LandingScreen"; // MyChartsScreen, PricingScreen, SubscriptionScreen, AllChartsScreen, DataSourceScreen
var tblDraggableBodyContent = "";
var tableOriginalBodyRows = "";
var myChartsSheetId = "";
var myChartsSheetName = "";
var syncMode = "Add";
var newlyAddedSheetNameOnEditMode = "MyChartSheet_b546-46b9-8b8a_CE";
var multipleUsersErrorDialogDisplayed = false;// dialog shown or not
var notAValidUserErrorDialogDisplayed = false;// dialog shown or not
var multipleUsersDiffLangErrorList = [];
var selectedDimensionMeasure = "";
var mychartsArray = "";
var currentOrder = "ascending";
var showInternetReconnectingOverlay = true;// for showing reconnecting overlay in case of internet disconnected
//var googleSheetV3LaunchDate = new Date("June 14, 2020 23:23:59");
var googleSheetV3LaunchDate = new Date("July 04, 2020 23:23:59");
var googleSheetTabularDataLaunchDate = new Date("February 10, 2021 05:00:01").getTime();
var samecontractlist = [];
var sheetRecords = [];
var datasourceColumnsWithIndex = [];
var dataSourceColumns = [[]];
var dataSourceColumnsAll = [[]];
var sheetA1NotationDetails = [];
var currentEditableCustomChartName = "";
multipleUsersDiffLangErrorList.push("required to perform that action");
multipleUsersDiffLangErrorList.push("Se necesita autorización para realizar esta acción");
multipleUsersDiffLangErrorList.push("Se requiere autorización para realizar esa acción");
multipleUsersDiffLangErrorList.push("Cal tenir autorització per efectuar aquesta acció");
multipleUsersDiffLangErrorList.push("Este necesară autorizarea pentru a efectua acțiunea respectivă");
multipleUsersDiffLangErrorList.push("autorização para efetuar");
multipleUsersDiffLangErrorList.push("autorização para executar");
multipleUsersDiffLangErrorList.push("Kailangan ng awtorisasyon upang maisagawa ang aksyon na iyan");
multipleUsersDiffLangErrorList.push("A művelet végrehajtásához engedély szükséges");
multipleUsersDiffLangErrorList.push("Da biste izvršili tu akciju, potrebna je autorizacija");
multipleUsersDiffLangErrorList.push("वह कार्यवाही करने के लिए अधिकार की आवश्यकता है");
multipleUsersDiffLangErrorList.push("ती क्रिया करण्यासाठी अधिकृतता आवश्यक आहे");
multipleUsersDiffLangErrorList.push("এই ক্রিয়াটি সম্পাদনা করার জন্য অনুমোদন প্রয়োজন৷");
multipleUsersDiffLangErrorList.push("Godkännande krävs för att utföra denna åtgärd");
multipleUsersDiffLangErrorList.push("Do wykonania tej czynności wymagana jest autoryzacja");
multipleUsersDiffLangErrorList.push("Vous devez disposer des autorisations requises pour pouvoir effectuer cette action");
multipleUsersDiffLangErrorList.push("Autorisation requise pour exécuter cette action. Exécutez à nouveau le script pour autoriser cette action");
multipleUsersDiffLangErrorList.push("Für die Ausführung dieser Aktion ist eine Berechtigung erforderlich");
multipleUsersDiffLangErrorList.push("K provedení dané akce je vyžadována autorizace");
multipleUsersDiffLangErrorList.push("Perlu otorisasi untuk melakukan tindakan itu");
multipleUsersDiffLangErrorList.push("За да извършите това действие, ви е необходимо разрешение");
multipleUsersDiffLangErrorList.push("その操作を実行するには承認が必要です");
multipleUsersDiffLangErrorList.push("அந்தச் செயலைச் செய்ய அங்கீகரிப்பு தேவைப்படுகிறது");
multipleUsersDiffLangErrorList.push("Для виконання цієї дії потрібно здійснити авторизацію");
multipleUsersDiffLangErrorList.push("Для выполнения этого действия необходима авторизация");
multipleUsersDiffLangErrorList.push("Autorisation er påkrævet");
multipleUsersDiffLangErrorList.push("richiesta l'autorizzazione");
multipleUsersDiffLangErrorList.push("toestemming nodig");
multipleUsersDiffLangErrorList.push("Bu eylemi gerçekleştirmek için yetki gerekiyor");
multipleUsersDiffLangErrorList.push("需要授權才能執行此動作");
multipleUsersDiffLangErrorList.push("Toiminnon tekemiseen vaaditaan lupa");
multipleUsersDiffLangErrorList.push("Cần được cho phép để thực hiện");

var maxNumberOfRecordsInSelectedSheet = 0;
var chartInEditModeSyncTime = false; // on new chart, it is false, but in case of editing existing chart from My Chart list, it will be true
var localStorageAccessible = false;
var storableObjectInTempStorage = {
    selectedChart: '',
    selectedChartDisplayName: '',
    editableChartCustomName: '',
    headerRow: '',
    dataRows: '',
    myChart: '',
    editableChartGuid: '',
    dimension: '',
    defaultProperties: '',
    selectedChartCategory: '',
    openDataViewerOnLoad: '',
    sameContractList: '',
    chartAddedUpdatedIntoMyChartList: '',
    headerRowNumber: '',
    headerRowAnnotation: '',
    dataRowsAnnotation: '',
    dataRowFrom: '',
    dataRowTo: '',
    fileName: '',
    fileId: '',
    sheetName: '',
    sheetId: '',
    isSampleData: '',
    sheetWholedataRows: '',
    propertiesSectionHeadsArray: '', // [] logically it will be array
    trialExpired: false,
    packageStatus: '',
    loggedInUserType: ''
};

var selectedchartDisplayName = "";
var syncSheetDataWithAddonTimeSpan = 10000; // 5 seconds
var syncSheetDataWithAddonTimerHandler = null;
var selectedChartNameFromSelectChartUI = "";
var selectedChartColumns = []; // contains selected (dim/metric)columns by user
var dataSourceDuplicateColumns = [];
var recentAction;// SHEET-DATA-NEW-CHART-ADDED, NEW-CHART
var chartGuid, clickedChartName, createdon;
var removeChartFromMyChartsListClickedFromMenuItem = false;
var chartRuleObject = {};

var dimensionColumns = [];
var metricColumns = [];

function saveDateInTempStorage(key, value) {
    if (localStorageAccessible) {
        window.localStorage.setItem(key, value);
    }
    else {
        storableObjectInTempStorage[key] = value;
    }
}

function getDateFromTempStorage(key) {
    if (localStorageAccessible) {
        return window.localStorage.getItem(key);
    }
    else {
        return storableObjectInTempStorage[key];
    }
}

//show current selectedSheet tooltip in myChart Screen
$('#addSheetClickContainer1').hover(function () {
    $('#addSheetClick1').attr('title', $(this).text());
}, function () {
    $('#addSheetClick1').attr('title', '');
});
$(document).ready(function () {
    $("#DataSourceDiv .containerBody").scroll(function () {
        if (selectedDimensionMeasure != "") {
            var leftPositionToMove = 40;
            var topPosition = $(selectedDimensionMeasure).position().top + $('#DataSourceDiv').scrollTop();
            var leftPosition = $(selectedDimensionMeasure).position().left;
            $("#myDropdown").css({ top: topPosition + 40, left: leftPosition < 100 ? leftPosition : leftPosition - leftPositionToMove });
        }
    });
    $('body').click(function () {
        $("#mainDropDownContainer").hide();
        if (selectedScreen == "SubscriptionScreen") {
            $('#tooltipSubscriptionScreen').html('');
        }
        $('#tooltipSubscriptionScreen').hide();
        hideSearchMenu();
        hideSearchSheetMenu();
        $('#addSheetClickContainer1').find('.imagecontainer').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    });
    // Divs for collapse/expand
    $("#divSankey").on("click", function () {
        collapseCategoryIcons();
        if ($("#divSankeyCharts").css("display") == "none") {
            $("#divSankeyCharts").css("display", "block");
            $("#divSentimentAnalysisCharts").css("display", "none");
            $("#divComparativeAnalysisCharts").css("display", "none");
            $("#divSpecializedSurveyCharts").css("display", "none");
            $("#divGeneralAnalysisCharts").css("display", "none");
            $("#divPPCCharts").css("display", "none");
            $('#imgSankey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $("#divSankeyCharts").css("display", "none");
            $('#imgSankey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
        }
    });

    //$(".myChartSelectSheetLabel").on("click", function () {
    //    $("#myDropdownSelectSheet").css("display", "block");
    //});

    $("#addSheetClickContainer1").on("click", function () {
        if ($("#myDropdownSelectSheet").css("display") == "none") {
            showSheetsSearchMenu(this);
            $("#myDropdownSelectSheet").css("display", "block");
            $('#addSheetClickContainer1').find('.imagecontainer').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $('#addSheetClickContainer1').find('.imagecontainer').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
            $("#myDropdownSelectSheet").css("display", "none");
        }
        clearFilterSheetFunction();
        $('#chartexpo_GoogleSheetAddon_tileMenu_txtSearch_selectSheet').focus();
        event.stopPropagation();

    });

    $(".topBarNotificationIcon").on("click", function () {
        logUserActionIntoDatabase("NotificationIconClicked", "Addon");
        ShowMainMenuViews('licenseKeyContainer');
        viewScreen("SubscriptionScreen");
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen != "SubscriptionScreen") {
            screenNavigationHistory.push("SubscriptionScreen");
        }
        if (loggedInUserCompleteDetail.LicencePackageDetail.packageStatus != "Purchased") {
            var subscriptionTypeVal = $("#radioBtnSingleUserReoccurringMonthlyPayment").attr("subscriptionType");
            if (subscriptionTypeVal == "Monthly") {
                $("#radioBtnSingleUserReoccurringMonthlyPayment").click();
            }
            else {
                $("#radioBtnSingleUserOneMonthPayment").click();
            }

            //if ($("#radioBtnSingleUserReoccurringMonthlyPayment").is(":checked")) {
            //    $(".btnIndividualAutoRenewSubscription").css("display", "table-cell");
            //    $("#btnIndividualUserBuyNow").css("display", "none");
            //    $("#radioBtnSingleUserReoccurringMonthlyPayment").attr('checked', true);
            //    $("#radioBtnSingleUserOneMonthPayment").attr('checked', false);
            //}
            //else {
            //    $(".btnIndividualAutoRenewSubscription").css("display", "none");
            //    $("#btnIndividualUserBuyNow").css("display", "table-cell");
            //    $("#radioBtnSingleUserReoccurringMonthlyPayment").attr('checked', false);
            //    $("#radioBtnSingleUserOneMonthPayment").attr('checked', true);
            //}
        }
        else {
            $("#btnUnsubscribeIndividualUser").show();
        }
    });

    $(document).on("click", "#btnManageDomain", function (e) {
        ShowMainMenuViews('licenseKeyContainer');
        viewScreen("SubscriptionScreen");
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen != "SubscriptionScreen") {
            screenNavigationHistory.push("SubscriptionScreen");
        }
        $('#tooltipSubscriptionScreen').hide();
        e.stopPropagation();
    });
    $(document).on("click", "#tooltipSubscriptionScreen", function () {
        $("#tooltipSubscriptionScreen").show();
    });
    $(".topBarNotificationIcon").hover(function () {
        if (selectedScreen != "SubscriptionScreen" && selectedScreen != "PricingScreen") {
            var html = $('#licenseKeyContainer').html();
            $('#tooltipSubscriptionScreen').html('');
            $('#tooltipSubscriptionScreen').append(html);
            $('.pageTitle').text('Subscribe');
            if (SubscriptionDomain == "Team") {
                $('.btnIndividualAutoRenewSubscription').hide();
                $('#btnIndividualUserBuyNow').hide();
                $('.separatorBorder').hide();
                $('.individualUserContainerSubscribeScreen').hide();
                $('#tooltipSubscriptionScreen').append('<div class="planSubscriptionSection"><div class="priceColumn subscriptionColumn"><p class="removeMargin_P"> $10 </p></div><div class="subscriptionColumn SubChargesColumn"><p class="subscriptionTextHead removeMargin_P">per user per month</p></div>');
                $('#tooltipSubscriptionScreen').append('<input id="btnManageDomain" type="button" value="Subscribe"/>');
                $('.teamUsersContainerSubscribeScreen').show();
                $('#tooltipSubscriptionScreen').find('.teamUsersContainerSubscribeScreen').remove();
            }
            if (loggedInUserCompleteDetail.LicencePackageDetail.packageStatus == "Purchased") {
                $("#btnUnsubscribeIndividualUser").show();
            }
            $('#tooltipSubscriptionScreen').show();
        }
    });
    $('#tooltipSubscriptionScreen').on("mouseleave", function () {
        if (SubscriptionDomain == "Individual") {
            var html = $('#tooltipSubscriptionScreen').html();
            $('#licenseKeyContainer').html('');
            $('#licenseKeyContainer').append(html);
        }
        else {
            $('.teamUsersContainerSubscribeScreen').show();
        }
        if (selectedScreen == "PricingScreen") {
            $('.pageTitle').text('ChartExpo Pricing');
        }
        else if (selectedScreen == "ManageTeamUsers") {
            $('.pageTitle').text('Manage Trial User(s)');
        }

        $('#tooltipSubscriptionScreen').hide();
    });
    $("#divSentimentAnalysis").on("click", function () {
        collapseCategoryIcons();
        if ($("#divSentimentAnalysisCharts").css("display") == "none") {
            $("#divSentimentAnalysisCharts").css("display", "block");
            $("#divSankeyCharts").css("display", "none");
            $("#divComparativeAnalysisCharts").css("display", "none");
            $("#divSpecializedSurveyCharts").css("display", "none");
            $("#divGeneralAnalysisCharts").css("display", "none");
            $("#divPPCCharts").css("display", "none");
            $('#imgSentimentAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $("#divSentimentAnalysisCharts").css("display", "none");
            $('#imgSentimentAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
        }
    });

    $("#divComparativeAnalysis").on("click", function () {
        collapseCategoryIcons();
        if ($("#divComparativeAnalysisCharts").css("display") == "none") {
            $("#divComparativeAnalysisCharts").css("display", "block");
            $("#divSankeyCharts").css("display", "none");
            $("#divSentimentAnalysisCharts").css("display", "none");
            $("#divSpecializedSurveyCharts").css("display", "none");
            $("#divGeneralAnalysisCharts").css("display", "none");
            $("#divPPCCharts").css("display", "none");
            $('#imgComparativeAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $("#divComparativeAnalysisCharts").css("display", "none");
            $('#imgComparativeAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
        }
    });

    $("#divSpecializedSurvey").on("click", function () {
        collapseCategoryIcons();
        if ($("#divSpecializedSurveyCharts").css("display") == "none") {
            $("#divSpecializedSurveyCharts").css("display", "block");
            $("#divSankeyCharts").css("display", "none");
            $("#divComparativeAnalysisCharts").css("display", "none");
            $("#divSentimentAnalysisCharts").css("display", "none");
            $("#divGeneralAnalysisCharts").css("display", "none");
            $("#divPPCCharts").css("display", "none");
            $('#imgSpecializedSurvey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $("#divSpecializedSurveyCharts").css("display", "none");
            $('#imgSpecializedSurvey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
        }
    });

    $("#divGeneralAnalysis").on("click", function () {
        collapseCategoryIcons();
        if ($("#divGeneralAnalysisCharts").css("display") == "none") {
            $("#divGeneralAnalysisCharts").css("display", "block");
            $("#divSankeyCharts").css("display", "none");
            $("#divSpecializedSurveyCharts").css("display", "none");
            $("#divComparativeAnalysisCharts").css("display", "none");
            $("#divSentimentAnalysisCharts").css("display", "none");
            $("#divPPCCharts").css("display", "none");
            $('#imgGeneralAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $("#divGeneralAnalysisCharts").css("display", "none");
            $('#imgGeneralAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
        }
    });

    $("#divPPC").on("click", function () {
        collapseCategoryIcons();
        if ($("#divPPCCharts").css("display") == "none") {
            $("#divPPCCharts").css("display", "block");
            $("#divSankeyCharts").css("display", "none");
            $("#divGeneralAnalysisCharts").css("display", "none");
            $("#divSpecializedSurveyCharts").css("display", "none");
            $("#divComparativeAnalysisCharts").css("display", "none");
            $("#divSentimentAnalysisCharts").css("display", "none");
            $('#imgPPC').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
        }
        else {
            $("#divPPCCharts").css("display", "none");
            $('#imgPPC').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
        }
    });

    $('#searchChartExpoCharts').on('input', function () {
        var re = new RegExp($(this).val(), "i"); // "i" means it's case-insensitive
        $('#divChartExpoCharts .chartsList .ChartListDivs').show().filter(function () {
            return !re.test($(this).text());
        }).hide();
        $('#divChartExpoCharts .chartsList .ChartListDivsWithScroll').show().filter(function () {
            return !re.test($(this).text());
        }).hide();
        searchChartExpoCharts();
    });

    $('#refreshChartExpoCharts').on('click', function () {
        $('#searchChartExpoCharts').val('');
        var re = new RegExp("", "i"); // "i" means it's case-insensitive
        $('#divChartExpoCharts .chartsList .ChartListDivs').show().filter(function () {
            return !re.test("");
        }).hide();
        $('#divChartExpoCharts .chartsList .ChartListDivsWithScroll').show().filter(function () {
            return !re.test("");
        }).hide();

        searchChartExpoCharts();
    });

    $('.ChartListDivs,.ChartListDivsWithScroll').on("click", function () {
        selectedChartNameFromSelectChartUI = $(this).attr('id');
        selectedChartCategory = $(this).attr('chartcategory');
        selectedChart = selectedChartNameFromSelectChartUI;
        initializeDataSourceScreenWithDefaultState();
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen != "DataSourceScreen") {
            screenNavigationHistory.push("DataSourceScreen");
        }
    });

    $(".maindropDownMenu_child").on("click", function () {
        var spanElementText = $(this).find("span").html();

        if (spanElementText == "My Chart") {
            if (isMySavedChartExists) {
                openMyChartsView();
            }
            var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
            if (lastScreen != "MyChartsScreen") {
                screenNavigationHistory.push("MyChartsScreen");
            }
        }
        else if (spanElementText == "Create New Chart") {
            if (trialExpired) {
                return;
            }
            else {
                showSelectChartScreen();
                viewScreen("AllChartsScreen");
                $('.subscriptionScreen').hide();
                $('.priceRangeDetailedScreen').hide();
                $('#DataSourceDiv').hide();
                boldSelectedTopMenuOption("normal", "normal", "normal", "normal", "Bold", "normal");
                var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
                if (lastScreen != "ChartListScreen") {
                    screenNavigationHistory.push("ChartListScreen");
                }
            }
        }
        else if (spanElementText == "Add Sample Chart Sheet") {
            if (trialExpired) {
                return;
            }
            else {
                boldSelectedTopMenuOption("normal", "normal", "normal", "normal", "normal", "Bold");
                insertChartSampleDataIntoSheet("Menu");
            }
        }
        else if (spanElementText == "Subscription") {
            clearTimeout(newchart_added_updated_timer_handler);
            if (trialExpired) {
                return;
                //showTrialTopBar(true, licenceMessage, -1, undefined, undefined, undefined);
                //ShowMainMenuViews('licenseKeyContainer');
                //viewScreen("SubscriptionScreen");
            }
            else {
                boldSelectedTopMenuOption("normal", "Bold", "normal", "normal", "normal", "normal");
                ShowMainMenuViews('licenseKeyContainer');
                viewScreen("SubscriptionScreen");
                logUserActionIntoDatabase("BuyNowScreenReviewed", "Addon");
                $(".selectedChartNameDiv").hide();
                var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
                if (lastScreen != "SubscriptionScreen") {
                    screenNavigationHistory.push("SubscriptionScreen");
                }
                //$(".buttonsContainer").hide();
                if (loggedInUserCompleteDetail.LicencePackageDetail.packageStatus != "Purchased") {
                    var subscriptionTypeVal = $("#radioBtnSingleUserReoccurringMonthlyPayment").attr("subscriptionType");
                    if (subscriptionTypeVal == "Monthly") {
                        $("#radioBtnSingleUserReoccurringMonthlyPayment").click();
                    }
                    else {
                        $("#radioBtnSingleUserOneMonthPayment").click();
                    }
                    //if ($("#radioBtnSingleUserReoccurringMonthlyPayment").is(":checked")) {
                    //    $(".btnIndividualAutoRenewSubscription").css("display", "table-cell");
                    //    $("#btnIndividualUserBuyNow").css("display", "none");
                    //    $("#radioBtnSingleUserReoccurringMonthlyPayment").attr('checked', true);
                    //    $("#radioBtnSingleUserOneMonthPayment").attr('checked', false);
                    //}
                    //else {
                    //    $(".btnIndividualAutoRenewSubscription").css("display", "none");
                    //    $("#btnIndividualUserBuyNow").css("display", "table-cell");
                    //    $("#radioBtnSingleUserReoccurringMonthlyPayment").attr('checked', false);
                    //    $("#radioBtnSingleUserOneMonthPayment").attr('checked', true);
                    //}
                }
                else {
                    $("#btnUnsubscribeIndividualUser").show();
                }
            }
        }
        else if (spanElementText == "Manage Domain") {
            boldSelectedTopMenuOption("normal", "normal", "normal", "Bold", "normal", "normal");
            clearTimeout(newchart_added_updated_timer_handler);

            if (trialExpired) {
                return;
                //showTrialTopBar(true, licenceMessage, -1, undefined, undefined, undefined);
                //ShowMainMenuViews('licenseKeyContainer');
                //viewScreen("SubscriptionScreen");
            }
            else {
                ShowMainMenuViews('licenseKeyContainer');
                viewScreen("SubscriptionScreen");
                logUserActionIntoDatabase("BuyNowScreenReviewed", "Addon");
                $(".selectedChartNameDiv").hide();
                //$(".buttonsContainer").hide();
                var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
                if (lastScreen != "SubscriptionScreen") {
                    screenNavigationHistory.push("SubscriptionScreen");
                }
            }
        }
        else if (spanElementText == "Help") {
            openHelpInNewTab();
            logUserActionIntoDatabase("ViewHelp-NavigationBar", "Addon");
        }
    });

    $("#divDimensionsContainer,#divMeasuresContainer,#DataSourceDiv").on('click', function () {
        hideSearchMenu();
    });

    $('#chkHeaderRow').change(function () {
        if ($(this).is(':checked')) {
            $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-active.png");
        }
        else {

            $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-deactive.png");
        }
        if ($("#addMeasureClick").css("display") == "none" || $("#addDimensionClick").css("display") == "none") {
            $("#addMeasureClick").css("display", "block");
            $("#addDimensionClick").css("display", "block");
        }
        $("#addMeasureClickContainer").css("margin-left", "0px");
        $("#addDimensionClickContainer").css("margin-left", "0px");
        onHeaderRowCheckboxChange(this);
        onHeaderChange();
    });

    // click event added to reload sheet dropdown on everytime user click on it to see items
    $('#dropdownSheets').on('click', function () {
        reloadSelectSheetDropdownListOnly();
    });

    $('#addDimensionClick').on('click', function () {
        $("#ulColumnsContainer").empty();
        clearTimeout(newchart_added_updated_timer_handler);
        if (isSampleDataClicked)
            return;

        if ($('#dropdownSheets').val() == "Select Sheet") {
            showMessageDialog("Select Sheet", "Please select any sheet.", "confirmation", false, [], true);
            return;
        }
        var chartRuleObject = getSelectedChartRulesObject();
        var dimensionText = 'dimension';
        var measureText = 'metric';
        if (chartRuleObject != null) {
            dimensionText = chartRuleObject.DimensionText;
            measureText = chartRuleObject.MeasureText;
        }
        if (dataSourceColumns[0].length == 0) {
            showMessageDialog("Select Sheet and Header Row", "Please select sheet and header row that have " + dimensionText + "(s) / " + measureText + "(s).", "confirmation", false, [], true);
            return;
        }
        var dimensionsInputLength = $('.dimension').length;
        var allowedDimensions = getSelectedChartAllowedDimensions();

        if (allowedDimensions == 0) {
            showMessageDialog("Rule Missing", "Please provide maximum allowed " + dimensionText + "(s) for this chart in rules!", "confirmation", false, [], true);
            return;
        }
        if (dimensionsInputLength >= allowedDimensions) {
            return;
        }
        logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-AddDimension_Clicked", "Charts");
        //showing search menu.
        isSearchMenuOpenedFromDimensionAddButton = true;
        isSearchMenuOpenedFromMeasureAddButton = false;
        //Reading header row on add dimension button click.
        getHeaderNewColumns(this);
    });

    $('#addMeasureClick').on('click', function () {
        //alert("addMeasureClick clicked");
        $("#ulColumnsContainer").empty();
        clearTimeout(newchart_added_updated_timer_handler);
        if (isSampleDataClicked)
            return;

        if ($('#dropdownSheets').val() == "Select Sheet") {
            //alert('Please select sheet');
            showMessageDialog("Select Sheet", "Please select any sheet.", "confirmation", false, [], true);
            return;
        }
        var chartRuleObject = getSelectedChartRulesObject();
        var dimensionText = 'dimension';
        var measureText = 'metric';
        if (chartRuleObject != null) {
            dimensionText = chartRuleObject.DimensionText;
            measureText = chartRuleObject.MeasureText;
        }
        if (dataSourceColumns[0].length == 0) {
            showMessageDialog("Select Sheet and Header Row", "Please select sheet and header row that have " + dimensionText + "(s) / " + measureText + "(s).", "confirmation", false, [], true);
            return;
        }
        var meausresInputLength = $('.metric').length;
        var allowedMetrics = getSelectedChartAllowedMeasures();

        if (allowedMetrics == 0) {
            showMessageDialog("Max Allowed Metric(s) Missing in Rules", "Please provide maximum allowed " + measureText + "(s) for this chart in rules!", "confirmation", false, [], true);
            return;
        }
        if (meausresInputLength >= allowedMetrics) {
            return;
        }
        logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-AddMeasure_Clicked", "Charts");
        //showing search menu.
        isSearchMenuOpenedFromMeasureAddButton = true;
        isSearchMenuOpenedFromDimensionAddButton = false;
        //Reading header row on add measure button click.
        getHeaderNewColumns(this);
    });

    $(".dimension").unbind("click");
    $(".dimension").on('click', function () {
        showSearchMenuOnDropdownClick(this, event);
    });
    $(".dropdownchangeclass").on('focus', function () {
        $(this).attr("disabled", "true");
        showSearchMenuOnDropdownClick(this.parentElement);
    });


    $("#chartexpo_GoogleSheetAddon_tileMenu").on('click', function () {
        event.stopPropagation();
    });


    $('#closeSearchColumnsDiv').click(function () {
        hideSearchMenu();
    });
    $('#addSampleSheetData').click(function () {
        insertChartSampleDataIntoSheet("Datasource");
    });
    $('#divTestDrawChart').click(function () {
        //let sheetName = "SankeyChart-SampleData";
        //var chartTitle = [["Sankey Chart Sample Data"]];
        //var tblHeader = [["Lead Source", "Lead Owner", "Duration", "Lead Status", "Count"]];
        //var tblRow = [["Advertisement", "Oliver", "Month or Less", "New", "73"], ["Customer event", "George", "3-6 Months", "New", "46"], ["Customer event", "Amelia", "3-6 Months", "Qualified", "73"], ["Employee referral", "Emily", "6-12 Months", "Working", "93"], ["Employee referral", "Amelia", "6-12 Months", "Qualified", "73"], ["Trade show", "Emily", "6-12 Months", "Working", "43"], ["Webinar", "Noah", "Over 1 Year", "Nurturing", "63"], ["Webinar", "Emily", "Over 1 Year", "Working", "41"], ["Website", "Noah", "Over 1 Year", "Nurturing", "46"], ["Other", "Jacob", "1-3 Months", "Unqualified", "39"], ["Other", "Isabella", "Over 1 Year", "Unqualified", "31"]];
        //var stepsHeader = [["How to create steps"]];
        //var stepsContent = [["1", "Select data source sheet"], ["2", "Set chart metric from your selected sheet data columns"], ["3", "Set sankey levels from selected sheet data columns"], ["4", "Select row range according to your requirement"], ["5", "Click on Create Chart button"], ["6", "For editing chart look, change its properties"]];
        //var dataHeader = "cell Inserted";
        //google.script.run.addSampleTestSheet(sheetName, JSON.stringify(chartTitle), JSON.stringify(tblHeader), JSON.stringify(tblRow), JSON.stringify(stepsHeader), JSON.stringify(stepsContent));
        var sheetName = SampleChartCreationSteps[selectedChart]();
        var timeStamp = new Date();
        var chartTitle = [[sheetName[0].chartName]];// + " SampleData"
        var headerColumns = [columnNameMapper[selectedChart]];
        var headerRows = [];
        var stepsHeader = [[sheetName[0].sheetHeaderTitle]];
        var stepsContent = sheetName[0].steps;
        var imagePath = "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/chartsampleimage/" + selectedChart + ".png";
        if (selectedChart == "ParetoGroupedChart" || selectedChart == "ParetoGroupedHorizontalChart") {
            imagePath = "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/chartsampleimage/" + selectedChart + ".png";
        }
        if (selectedChartCategory !== undefined && selectedChartCategory === "PPC") {
            headerRows = PPCChartsSampleData[selectedChart]();
        }
        else {
            headerRows = SampleData[selectedChart]();
        }
        var tblHeader = [];
        var tblBody = [];
        if (selectedChart == "SentimentTrendChart") {
            tblBody = sentimentTrendGetDataRowForExcel(headerRows);
        }
        else {
            // for multi measure charts call MultiMeasure method
            if (selectedChart == "GaugeChart" || selectedChart == "ISGraph" || selectedChart == "ScatterChartAdvance" ||
                selectedChart == "HierarichalBarChartAdvance" || selectedChart == "ChordChart" || selectedChart == "DoubleMeasureComparisonChart") {
                if (selectedChart == "GaugeChart") {
                    var gaugeData = [];
                    gaugeData.push(headerRows)
                    data = gaugeData;
                }
                tblBody = getDataRowForExcelMultiMeasure(headerRows);
            }
            else {
                tblBody = getDataRowForExcel(headerRows);
            }
        }
        //convert data object into 2d array
        for (var i = 0; i < headerColumns.length; i++) {
            tblHeader.push(headerColumns[i]);
        }
        google.script.run.addSampleTestSheet(sheetName[0].chartName + "-SampleData-" + timeStamp.getTime(), JSON.stringify(chartTitle), JSON.stringify(tblHeader), JSON.stringify(tblBody), JSON.stringify(stepsHeader), JSON.stringify(stepsContent), imagePath, sheetName[0].chartName);
    });

    $('.selectedChartIcon').click(function () {

        //$('#DataSourceDivHeaderRow').hide();
        $('#DataSourceDiv').hide();
        $('.selectedChartNameDiv').hide();
        $("#divChartExpoChartsSearchBox").show();
        $("#divChartExpoCharts").show();
        selectedScreen = "AllChartsScreen";
        $('.samplesheetdata').hide();
        //$('#divSankeyCharts').show();
        $('.rearrangeText').css("z-index", "100");
        $("input").focus();
        if (enableCurrentActiveSheetSelection) {
            //When move from DataSource and click again syncMode changed to Edit instead of Add so manually change it.
            syncMode = "Add";
        }
        screenNavigationHistory.pop();
    });


    $(document).on("mouseenter", ".dimension", function (event) {
        if (!isSampleDataClicked) {
            if ($(this).find('div').attr('isdeletedcolumn') == "false") {
                $(this).find('.dropdowncolor').css("background-color", "#FDE4D6");
                $(this).find('.DropdownList').attr('title', $(this).text());
            }
            else {
                var columnName = $(this).text();
                columnName = $.trim(columnName);
                var html = '<span>' +
                    'The column does not exist. Please map it with a valid sheet column.' +
                    '</span>';
                generateTooltip(false, "deletedDimension", html, event, this);
            }
            $(this).find('.dimensionRemoveClass').css("visibility", "visible");
        }
    });
    $(document).on("mouseleave", ".dimension", function () {
        if (!isSampleDataClicked) {
            if ($(this).find('div').attr('isdeletedcolumn') == "false") {
                $(this).find('.dropdowncolor').css("background-color", "white");
                $(this).find('.DropdownList').attr('title', '');
            }
            else {
                hideTooltip();
            }
            $(this).find('.dimensionRemoveClass').css("visibility", "hidden");
        }
    });
    $(document).on("mouseenter", ".metric", function (event) {
        if (!isSampleDataClicked) {
            if ($(this).find('div').attr('isdeletedcolumn') == "false") {
                $(this).find('.metricdropdowncolor').css("background-color", "#FDE4D6");
                $(this).find('.DropdownList').attr('title', $(this).text());
            }
            else {
                var columnName = $(this).text();
                columnName = $.trim(columnName);
                var html = '<span>' +
                    'The column does not exist. Please map it with a valid sheet column.' +
                    '</span>';
                generateTooltip(false, "deletedMetric", html, event, this);
            }
            $(this).find('.metricRemoveClass').css("visibility", "visible");
        }
    });
    $(document).on("mouseleave", ".metric", function () {
        if (!isSampleDataClicked) {
            if ($(this).find('div').attr('isdeletedcolumn') == "false") {
                $(this).find('.metricdropdowncolor').css("background-color", "white");
                $(this).find('.DropdownList').attr('title', '');
            }
            else {
                hideTooltip();
            }
            $(this).find('.metricRemoveClass').css("visibility", "hidden");
        }
    });

    $('#startRowTextBox').on('change', function () {
        logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-StartRangeTextBoxChanged", "Charts");
        var max = parseInt($(this).attr('max'));
        var min = parseInt($(this).attr('min'));
        var value = +$(this).val();
        if (value > max) {
            $(this).val(max);
            value = max;
        }
        else if (value < min) {
            $(this).val(min);
            value = min;
        }
        else if (value > +$("#endRowTextBox").val()) {
            $(this).val(min);
            value = min;
        }

        if ($("#chkHeaderRow").prop("checked")) {
            updateSliderValue(value - 1, +$("#endRowTextBox").val() - 1);
        }
        else {
            updateSliderValue(value, +$("#endRowTextBox").val());
        }
        // console.log("syncSelectedChartAtServer called from $('#startRowTextBox').on('change' ");
        if (!isSampleDataClicked) {
            syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
        }
    });

    $('#endRowTextBox').on('change', function () {
        logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-EndRangeTextBoxChanged", "Charts");
        var max = parseInt($(this).attr('max'));
        var min = parseInt($(this).attr('min'));
        var value = +$(this).val();
        if (value > max) {
            $(this).val(max);
            value = max;
        }
        else if (value < min) {
            $(this).val(min);
            value = min;
        }
        else if (value < +$("#startRowTextBox").val()) {
            $(this).val(max);
            value = max;
        }

        if ($("#chkHeaderRow").prop("checked")) {
            updateSliderValue(+$("#startRowTextBox").val() - 1, value - 1);
        }
        else {
            updateSliderValue(+$("#startRowTextBox").val(), value);
        }

        // console.log("syncSelectedChartAtServer called from $('#endRowTextBox').on('change' ");
        if (!isSampleDataClicked) {
            syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
        }
    });


    $('#info-icon-dimension').hover(function (event) {
        var chartRuleObject = getSelectedChartRulesObject();
        var html;
        var column = "a column";
        var column1 = "column";
        var displayName = chartRuleObject.ChartDisplayName;
        var dimension = chartRuleObject.DimensionText;
        var only = "only ";
        var rearrangeColumns = "";
        if (chartRuleObject.MaxDim > 1) {
            column = "columns";
            column1 = "columns";
            dimension = chartRuleObject.DimensionText + "s";
            rearrangeColumns = '<br><br> You can rearrange ' + dimension.toLowerCase() + ' with a drag-n-drop';
            only = "";
        }
        if (chartRuleObject.ChartDisplayName == "Sankey Chart" || chartRuleObject.ChartDisplayName == "Sankey Aggregated Chart" || chartRuleObject.ChartDisplayName == "Sankey Sentiment Chart") {
            displayName = displayName.replace("Chart", "");
        }
        if (chartRuleObject.MinDim == chartRuleObject.MaxDim) {
            html = '<span>' +
                'Select ' + column + ' for ' + displayName + ' ' + dimension + ' that contains categorical data (e.g. device type, geo, etc.)' +
                '<br><br>Don\'t select ' + column + ' that contains numeric data (e.g. clicks, cost, etc.)' +
                '<br><br>You can select ' + only + chartRuleObject.MinDim + ' ' + column1 +
                rearrangeColumns +
                '</span>';
        }
        else {
            html = '<span>' +
                'Select ' + column + ' for ' + displayName + ' ' + dimension + ' that contains categorical data (e.g. device type, geo, etc.)' +
                '<br><br>Don\'t select ' + column + ' that contains numeric data (e.g. clicks, cost, etc.)' +
                '<br><br>You can select between ' + chartRuleObject.MinDim + ' and ' + chartRuleObject.MaxDim + ' ' + column +
                rearrangeColumns +
                '</span>';
        }
        generateTooltip(false, "dimensions", html, event);
    }, function () {
        hideTooltip();
    });
    $('#info-icon-metric').hover(function (event) {
        var chartRuleObject = getSelectedChartRulesObject();
        var html;
        var column = "a column";
        var column1 = "column";
        var metric = chartRuleObject.MeasureText;
        var displayName = chartRuleObject.ChartDisplayName;
        var only = "only ";
        var rearrangeColumns = '';
        if (chartRuleObject.MaxMetric > 1) {
            column = "columns";
            column1 = "columns";
            metric = chartRuleObject.MeasureText + "s";
            rearrangeColumns = '<br><br> You can rearrange ' + metric.toLowerCase() + ' with a drag-n-drop';
            only = "";
        }
        if (chartRuleObject.ChartDisplayName == "Sankey Chart" || chartRuleObject.ChartDisplayName == "Sankey Aggregated Chart" || chartRuleObject.ChartDisplayName == "Sankey Sentiment Chart") {
            displayName = displayName.replace("Chart", "");
        }
        if (chartRuleObject.MinMetric == chartRuleObject.MaxMetric) {
            html = '<span>' +
                'Select ' + column + ' for ' + displayName + ' ' + metric + ' that contains numeric data (e.g. clicks, cost, etc.)' +
                '<br><br>Don\'t select ' + column + ' that contains categorical data (e.g. device type, geo, etc.)' +
                '<br><br>You can select ' + only + chartRuleObject.MinMetric + ' ' + column1 +
                rearrangeColumns +
                '</span>';
        }
        else if (chartRuleObject.MinMetric < chartRuleObject.MaxMetric && chartRuleObject.MaxMetric == 2)
        {
            html = '<span>' +
                'Select ' + column + ' for ' + displayName + ' ' + metric + ' that contains numeric data (e.g. clicks, cost, etc.)' +
                '<br><br>Don\'t select ' + column + ' that contains categorical data (e.g. device type, geo, etc.)' +
                '<br><br>You can select ' + chartRuleObject.MinMetric + ' or ' + chartRuleObject.MaxMetric + ' ' + column +
                rearrangeColumns +
                '</span>';
        }
        else {
            html = '<span>' +
                'Select ' + column + ' for ' + displayName + ' ' + metric + ' that contains numeric data (e.g. clicks, cost, etc.)' +
                '<br><br>Don\'t select ' + column + ' that contains categorical data (e.g. device type, geo, etc.)' +
                '<br><br>You can select between ' + chartRuleObject.MinMetric + ' and ' + chartRuleObject.MaxMetric + ' ' + column +
                rearrangeColumns +
                '</span>';
        }
        generateTooltip(false, "metrics", html, event);
    }, function () {
        hideTooltip();
    });
    $('.getHelpContactUsDiv').hover(function (event) {
        var html = '<span>' +
            'Send us an email, we will do our best to help you.' +
            '</span>';
        generateTooltip(false, "help", html, event, this);
    }, function () {
        hideTooltip();
    });
    $('#headerRowLabel').hover(function (event) {
        var html = '<span>' +
            'Uncheck the box if you do not have column header. Column names will be used instead.' +
            '</span>';
        if ($('#btnDrawChartFromSheetData').hasClass('tabButtonActive')) {
            var x = $('#headerRowLabel').offset().left;
            var y = $('#headerRowLabel').offset().top;
            var arrowheight = 8;
            y = y + $('#headerRowLabel').height() + arrowheight;
            x = x - 6;
            generateTooltip(false, "headerRow", html, event);
            $('.info .toolTip-headerRow').css('top', y);
            $('.info .toolTip-headerRow').css('left', x);
        }
    }, function () {
        hideTooltip();
    });
    $('#addSampleSheetData').hover(function (event) {
        var html = '<span>' +
            'Click here to add a sheet with sample data and chart.' +
            '</span>';
        var x = $('#addSampleSheetData').offset().left;
        var y = $('#addSampleSheetData').offset().top;
        var arrowheight = 8;
        y = y + $('#addSampleSheetData').height() + arrowheight;
        x = x - 6;
        generateTooltip(false, "addSheetSampleData", html, event);
        $('.info .toolTip-sheetSampleData').css('top', y);
        $('.info .toolTip-sheetSampleData').css('left', x);
    }, function () {
        hideTooltip();
    });
    $('#dropdownSheets').hover(function () {
        $(this).attr("title", $("#dropdownSheets option:selected").val());
    }, function () {
        $(this).attr("title", "");
    });
    $('#divDrawChart').hover(
        function () {
            if (isSampleDataClicked) {
                $('#divDrawChart').css({
                    "background-color": "#F37A2D", "color": "white"
                });
            }
            else {
                // activeDrawButton class only applied if, chart data source screen has all required dimensions and measures
                if ($('#divDrawChart').hasClass("activeDrawButton") == false) {
                    var chartRuleObject = getSelectedChartRulesObject();
                    showTooltipOnCreateChartButton(chartRuleObject, this);
                }
                else {
                    $('#divDrawChart').css({
                        "background-color": "#F37A2D", "color": "white"
                    });
                }
            }
        },
        function () {
            if (isSampleDataClicked) {
                $('#divDrawChart').css({
                    "background-color": "white", "color": "#F37A2D"
                });
            }
            else {
                if ($('#divDrawChart').hasClass("activeDrawButton") == false) {
                    //var chartRuleObject = getSelectedChartRulesObject();
                    //showTooltipOnCreateChartButton(chartRuleObject);

                    hideTooltip();
                }
                else {
                    $('#divDrawChart').css({
                        "background-color": "white", "color": "#F37A2D"
                    });
                }
            }
        }
    );

    $('#divDrawChart').on('click', function (event) {
        var chartRuleObject = getSelectedChartRulesObject();
        var selectedSheetName = $('#dropdownSheets').val();
        if (isSampleDataClicked) {
            $(".se-pre-con").fadeIn("slow");
            saveDateInTempStorage("isSampleData", "true");
            createChartWithLatestData();
        }
        else {
            var chartRuleObject = getSelectedChartRulesObject();
            saveDateInTempStorage("isSampleData", "false");
            if ($('#dropdownSheets').val() == "Select Sheet") {
                //showMessageDialog("Select Sheet", "Please select any sheet.", "confirmation", false, [], true);
                showTooltipOnCreateChartButton(chartRuleObject, this);
                return;
            }

            var dimensions = [];
            var measures = [];
            if (chartRuleObject != null) {
                var isDeletedColumnExist = false;
                $(".dimension > div:first-child").each(function (index) {
                    var isColumnDeleted = $(this).attr('isdeletedcolumn');
                    if (isColumnDeleted == "true") {
                        isDeletedColumnExist = true;
                    }
                    dimensions.push($(this).text());
                });
                if (isDeletedColumnExist) {
                    return;
                }
                $(".metric > div:first-child").each(function (index) {
                    var isColumnDeleted = $(this).attr('isdeletedcolumn');
                    if (isColumnDeleted == "true") {
                        isDeletedColumnExist = true;
                    }
                    measures.push($(this).text());
                });
                if (isDeletedColumnExist) {
                    return;
                }

                if (dimensions.length < chartRuleObject.MinDim || measures.length < chartRuleObject.MinMetric) {
                    var dimText = " " + chartRuleObject.DimensionText.toLowerCase();
                    var measureText = " " + chartRuleObject.MeasureText.toLowerCase() + ".";
                    if (chartRuleObject.MinDim > 1) {
                        dimText = " " + chartRuleObject.DimensionText.toLowerCase() + "s";
                    }
                    if (chartRuleObject.MinMetric > 1) {
                        measureText = " " + chartRuleObject.MeasureText.toLowerCase() + "s.";
                    }
                    //var message = "Please select minimum " + chartRuleObject.MinDim + dimText + " and " + chartRuleObject.MinMetric + measureText;
                    //showMessageDialog("Select " + chartRuleObject.DimensionText + " and " + chartRuleObject.MeasureText + "", message, "confirmation", false, [], true);

                    showTooltipOnCreateChartButton(chartRuleObject);

                    return;
                }
                samecontractlist = getSameContractCharts(chartRuleObject.MinDim, chartRuleObject.MinMetric, chartRuleObject);
            }

            // First we will get latest A1Notation from sheet and then set other variables state for chart viewer
            //
            getLatestA1NotationFromSheet(selectedSheetName, this);
        }
        //window.localStorage.setItem("chartAddedUpdatedIntoMyChartList", "0");

        saveDateInTempStorage("chartAddedUpdatedIntoMyChartList", "0");

        newchart_added_updated_timer_handler = setTimeout(chartAddedUpdatedInMyCharts, timeSpaneToCheckMyChart);
    });

    $('.CreateChartDiv').click(function () {
        chartInEditModeSyncTime = false;
        if (trialExpired) {
            return;
            //showTrialTopBar(true, licenceMessage, -1, undefined, undefined, undefined);
            //ShowMainMenuViews('licenseKeyContainer');
            //viewScreen("SubscriptionScreen");
        }
        else {
            syncMode = "Add";
            recentAction = "NEW-CHART";
            $('#CreateChartContentDiv').hide();
            $('#ChartOptionsDiv').hide();
            logUserActionIntoDatabase("ViewChartsList-LandingPage", "Addon");
            resetSearchChartOnNewAction();
            showSelectChartScreen();
            viewScreen("AllChartsScreen");
            var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
            if (lastScreen != "ChartListScreen") {
                screenNavigationHistory.push("ChartListScreen");
            }
        }
    });

    $('.ChartDivs').click(function () {
        $('#CreateChartContentDiv').hide();
        $('#DropdownChartsDiv').hide();
        $('#ChartOptionsDiv').show();
    });

    $('.FloatRightEditChart').click(function () {

        $('.ChartdimensionContainer').show();
        $('#SampleDataTableDiv').hide();
    });

    $('.FloatRightPropertySampleData').click(function () {
        $('.ChartdimensionContainer').hide();
        $('#SampleDataTableDiv').show();
    });

    $('#CopytxtClipboard').click(function () {
        CopyToClipboard();
    });

    //$('#addchartIcon').click(function () {
    //    chartInEditModeSyncTime = false;
    //    if (trialExpired) {
    //        return;
    //        //showTrialTopBar(true, licenceMessage, -1, undefined, undefined, undefined);
    //        //ShowMainMenuViews('licenseKeyContainer');
    //        //viewScreen("SubscriptionScreen");
    //    }
    //    else {
    //        syncMode = "Add";
    //        recentAction = "NEW-CHART";
    //        logUserActionIntoDatabase("ViewChartsList-LandingPage", "Addon");

    //        resetSearchChartOnNewAction();

    //        showSelectChartScreen();
    //        viewScreen("AllChartsScreen");
    //    }
    //});

    $('#myChartsSearchBox').on('input', function () {
        var re = new RegExp($(this).val(), "i"); // "i" means it's case-insensitive
        /*
        $('.thumbnail .thumbailChartTitle .thumbailChartTitlePart').show().filter(function () {
        console.log($(this).val());
    return !re.test($(this).text());
        }).hide();*/

        $('.RowOptionsDiv').show().filter(function () {
            return !re.test($(this).text());
        }).hide();
    });


    $("#myChartsThumbnailsDiv").on("click", ".openMySavedChart", function () {
        var isconnected = checkInternetConnection();
        if (!isconnected) {
            $(".se-pre-con").fadeOut("slow");
            showReconnectingOverlay();
            checkInternetHandler = setInterval(checkInternetConnection, 3000);
            return;
        }
        var chartGuid = $(this).attr("chartGuid");
        var clickedChartName = $(this).attr("charttype");
        var createdon = $(this).attr("createdon");
        currentEditableCustomChartName = $(this).attr('editableChartName');
        synchedChartGUID = chartGuid; // for syncMode
        syncMode = "Edit";
        $(".se-pre-con").fadeIn("slow");
        logUserActionIntoDatabase(clickedChartName + "-OpenedFromMyChartsList", "Charts");

        google.script.run.withSuccessHandler(function (addonInitialStateObj) {
            var myChartMeta = JSON.parse(addonInitialStateObj);

            selectedChart = mySavedListSelectedChartName = myChartMeta.ChartName;
            var savedSettings = myChartMeta.ChartMetaJSON;
            var savedHeaderRow = null, chartDimension = {};
            var useHeaderRow = "true";
            var a1NotationInformation = null;
            var dataRowFrom, dataRowTo, headerRowNumber, headerRowAnnotation, dataRowsAnnotation, selectedDimensions, selectedMeasures;

            mySavedListSelectedChartCategory = undefined;

            var editableMyChartCompleteDetail = {};

            // hanlde old saved my chart as need to show chart header count as per chart data
            mySavedListSelectedChartData = JSON.parse(myChartMeta.ChartDataJSON);

            /*
            // get detail of the node to be
            var chartAddedIntoMyChartListDate = new Date(+createdon); // createdOn

            if (chartAddedIntoMyChartListDate < googleSheetV3LaunchDate) {
        editableMyChartCompleteDetail.dataRows = getSampleDataRows(selectedChart, mySavedListSelectedChartData);
    }
            else {
        editableMyChartCompleteDetail.dataRows = mySavedListSelectedChartData;
    }
            */
            editableMyChartCompleteDetail.dataRows = mySavedListSelectedChartData;

            if (savedSettings != undefined) {
                chartInEditModeSyncTime = true;
                savedSettings = JSON.parse(savedSettings);
                mySavedListSelectedChartCategory = savedSettings.chartCategory;
                mySavedListSelectedChartProperties = savedSettings.props;
                chartDimension = savedSettings.dimension;

                if (savedSettings.headerRow != undefined) {
                    savedHeaderRow = JSON.parse(savedSettings.headerRow);
                }
                else {
                    if (editableMyChartCompleteDetail.dataRows.length > 0) {
                        savedHeaderRow = getSampleDataHeaderRow(selectedChart, editableMyChartCompleteDetail.dataRows[0].length);
                    }
                    else {
                        savedHeaderRow = getSampleDataHeaderRow(selectedChart);
                    }
                }

                if (mySavedListSelectedChartProperties != undefined) {
                    mySavedListSelectedChartProperties = JSON.stringify(mySavedListSelectedChartProperties).replace("___br___", "<br>");
                }

                if (savedSettings.selectedDimensions != undefined) {
                    selectedDimensions = JSON.parse(savedSettings.selectedDimensions);
                }
                if (savedSettings.selectedMeasures != undefined) {
                    selectedMeasures = JSON.parse(savedSettings.selectedMeasures);
                }

                a1NotationInformation = savedSettings.a1NotationInformation;

                useHeaderRow = savedSettings.useHeaderRow;
                fileName = savedSettings.fileName;
                sheetName = savedSettings.sheetName;
                fileId = savedSettings.fileId;
                sheetId = savedSettings.sheetId;
                headerRowAnnotation = savedSettings.headerRowAnnotation;
                dataRowsAnnotation = savedSettings.dataRowsAnnotation;
                headerRowNumber = savedSettings.headerRowNumber;
                dataRowFrom = savedSettings.dataRowFrom;
                dataRowTo = savedSettings.dataRowTo;
            }

            var chartRuleObject = getSelectedChartDisplayName();

            selectedchartDisplayName = chartRuleObject.ChartDisplayName;
            editableMyChartCompleteDetail.selectedChart = selectedChart;
            editableMyChartCompleteDetail.selectedChartDisplayName = selectedchartDisplayName;
            editableMyChartCompleteDetail.selectedChartCategory = mySavedListSelectedChartCategory;

            editableMyChartCompleteDetail.headerRow = savedHeaderRow;

            editableMyChartCompleteDetail.a1NotationInformation = a1NotationInformation;

            editableMyChartCompleteDetail.useHeaderRow = useHeaderRow;
            editableMyChartCompleteDetail.headerRowAnnotation = headerRowAnnotation;
            editableMyChartCompleteDetail.dataRowsAnnotation = dataRowsAnnotation;
            editableMyChartCompleteDetail.fileName = fileName;
            editableMyChartCompleteDetail.sheetName = sheetName;
            editableMyChartCompleteDetail.fileId = fileId;
            editableMyChartCompleteDetail.sheetId = sheetId;
            editableMyChartCompleteDetail.headerRowNumber = headerRowNumber;
            editableMyChartCompleteDetail.dataRowFrom = dataRowFrom;
            editableMyChartCompleteDetail.dataRowTo = dataRowTo;

            editableMyChartCompleteDetail.selectedDimensions = selectedDimensions;
            editableMyChartCompleteDetail.selectedMeasures = selectedMeasures;

            editableMyChartCompleteDetail.defaultProperties = mySavedListSelectedChartProperties;
            synchedChartProperties = JSON.parse(mySavedListSelectedChartProperties); // used for next processing in case of sync mode
            synchedChartDimensions = chartDimension;

            editableMyChartCompleteDetail.myCharts = "true";
            editableMyChartCompleteDetail.dimension = chartDimension;
            editableMyChartCompleteDetail.editableChartGuid = chartGuid;

            samecontractlist = getSameContractCharts(chartRuleObject.MinDim, chartRuleObject.MinMetric, chartRuleObject);

            storeDataInLocalStorage(editableMyChartCompleteDetail);
            saveDateInTempStorage("isSampleData", "false"); // this flag used to set auto synching in chart viewer, if true do not sync, otherwise sync
            openChartViewer("mycharts", editableMyChartCompleteDetail.headerRow, editableMyChartCompleteDetail.dataRows, selectedChart);
        })
            .withFailureHandler(
            function (msg, element) {
                $(".se-pre-con").fadeOut("slow");
                handleError(msg);
            }
            ).getSelectedChartMeta(chartGuid);
    });

    $("#myChartsThumbnailsDiv").on("click", ".removeFromMySavedCharts", function () {

        var chartGuid = $(this).attr("chartguid");
        var clickedChartName = $(this).attr("charttype");
        removableChartGuid = chartGuid;
        removableMyChartName = clickedChartName;
        showMessageDialog("Remove Chart", "Are you sure to remove chart from My Chart list.", "confirmation", true, ["Yes", "Cancel"], false);
        //removeChartFromMyChartsLis(chartGuid);
    });

    // Code to run on Edit my saved chart
    $("#myChartsThumbnailsDiv").on("click", ".editMySavedChart", function () {
        var isconnected = checkInternetConnection();
        if (!isconnected) {
            $(".se-pre-con").fadeOut("slow");
            showReconnectingOverlay();
            checkInternetHandler = setInterval(checkInternetConnection, 3000);
            return;
        }
        viewScreen("DataSourceScreen");
        $(".se-pre-con").fadeIn("slow");
        //$(".buttonsContainer").show();
        var chartGuid = $(this).attr("chartGuid");
        var clickedChartName = $(this).attr("charttype");
        var createdon = $(this).attr("createdon");
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen != "DataSourceScreen") {
            screenNavigationHistory.push("DataSourceScreen");
        }
        currentEditableCustomChartName = $(this).attr('editableChartName');
        synchedChartGUID = chartGuid; // for syncMode
        syncMode = "Edit";

        //openSelectedChartGuidInEditMode(chartGuid);
        openSelectedChartGuidInEditMode(chartGuid);
    });

    $("#myChartsThumbnailsDiv").on("click", ".CreateNewFromSavedChart", function () {
        var isconnected = checkInternetConnection();
        if (!isconnected) {
            $(".se-pre-con").fadeOut("slow");
            showReconnectingOverlay();
            checkInternetHandler = setInterval(checkInternetConnection, 3000);
            return;
        }
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen != "DataSourceScreen") {
            screenNavigationHistory.push("DataSourceScreen");
        }
        var clickedChartName = $(this).attr("charttype");
        selectedChartNameFromSelectChartUI = clickedChartName;
        selectedChart = selectedChartNameFromSelectChartUI;
        selectedChartCategory = $(this).attr('chartcategory');
        initializeDataSourceScreenWithDefaultState();
    });

    $("#myChartsThumbnailsDiv").on("click", ".insertChartIntoSheet", function () {
        //console.log("load script dynamilcally");

        chartGuid = $(this).attr("chartGuid");
        clickedChartName = $(this).attr("charttype");
        createdon = $(this).attr("createdon");
        selectedChart = clickedChartName;

        AddScript("https://doc-04-ao-docs.googleusercontent.com/docs/securesc/4rk5cof9mh9g2noe3od806r4a21a2tb9/3afilur1j7n84qh6uvjqkfo0ceiq57qv/1619440800000/10040362025876787199/10040362025876787199/1EklpkOGdEzqX8jn1fnUQncQwVBH_HOmQ?e=download&authuser=0&nonce=0ru7594u98hjm&user=10040362025876787199&hash=4p9pg90kqf5h9fpoocokrgomep4cqim6", "ChartExpo", runInsertChartImageIntoSheetMethod, true);
        //        AddScript("https://chartexpo.com/ChartExpoForGoogleSheetAddin/Scripts/Polyvista/ChartExpo.February.v1221.js", "ChartExpo", runInsertChartImageIntoSheetMethod, true);
        /*
            //
            var chartGuid = $(this).attr("chartGuid");
            var clickedChartName = $(this).attr("charttype");
            var createdon = $(this).attr("createdon");
            selectedChart = clickedChartName;
            $(".se-pre-con").fadeIn("slow");
            logUserActionIntoDatabase(clickedChartName + "-Chart_Image_Inserted", "Charts");
            // Get this chart
            google.script.run.withSuccessHandler(function (chartMetaObject) {
                var myChartMeta = JSON.parse(chartMetaObject);
                var chartAddedIntoMyChartListDate = new Date(+createdon); // createdOn
                // TODO, need to write its conversion code
                if (chartAddedIntoMyChartListDate < googleSheetV3LaunchDate) {
        myChartMeta.ChartDataJSON = JSON.parse(myChartMeta.ChartDataJSON);//getSampleDataRows(clickedChartName, myChartMeta.ChartDataJSON); 
    }
                else {
        myChartMeta.ChartDataJSON = convertDataIntoChartFormat(JSON.parse(myChartMeta.ChartDataJSON));
    }
                //console.log(JSON.stringify(myChartMeta.ChartDataJSON));
                drawChartForImage(myChartMeta);
            })
                .withFailureHandler(
                function (msg, element) {
        $(".se-pre-con").fadeOut("slow");
    alert(msg);
                }
                ).getSelectedChartMeta(chartGuid);
                */
    });

    // Code to run on click on my chart new option 'Add new sample data and sheet'
    $("#myChartsThumbnailsDiv").on("click", ".addSampleSheet", function () {
        selectedChart = $(this).attr("charttype");
        selectedChartCategory = $(this).attr("chartCategory");
        selectedChartNameFromSelectChartUI = selectedChart;
        syncMode = "Add";
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen != "DataSourceScreen") {
            screenNavigationHistory.push("DataSourceScreen");
        }
        logUserActionIntoDatabase(selectedChart + "-SampleSheetInserted", "Charts");
        insertChartSampleDataIntoSheet("MyChart");
        viewScreen("DataSourceScreen");
    });

    // call logUserActionIntoDatabase(selectedChart + "-SampleSheetInserted", "Charts"); in case of sample data clicked from menu as well

    $(".error_dialog_button_cancel").on("click", function () {
        hideMessageDialog();
        if (multipleUsersErrorDialogDisplayed != undefined && multipleUsersErrorDialogDisplayed) {
            google.script.host.close();
        }

        if (notAValidUserErrorDialogDisplayed != undefined && notAValidUserErrorDialogDisplayed) {
            google.script.host.close();
        }
    });

    $(".error_dialog_closeButton").click(function (e) {
        hideMessageDialog();
        if (multipleUsersErrorDialogDisplayed != undefined && multipleUsersErrorDialogDisplayed) {
            google.script.host.close();
        }

        if (notAValidUserErrorDialogDisplayed != undefined && notAValidUserErrorDialogDisplayed) {
            google.script.host.close();
        }
    });

    $(".multipleLoginMessageCloseLink").click(function (e) {
        google.script.host.close();
    });

    $(".error_dialog_button_view").on("click", function () {
        if ($(this).val() === "Yes") {
            // call chart removal code
            hideMessageDialog();
            logUserActionIntoDatabase(removableMyChartName + "-RemovedFromMyChartsList", "Charts");
            removeChartFromMyChartsLis(removableChartGuid);
        }
        else if ($(this).val() === "View") {
            // call view charts list method
            hideMessageDialog();
            openMyChartsContainerView();//'ViewMyCharts-NewChartAdded');
        }
    });

    //$("#btnDrawChartFromSampleData").on("click", function () {
    //    enableControl();
    //    clearTimeout(newchart_added_updated_timer_handler);
    //    clearTimeout(syncSheetDataWithAddonTimerHandler);
    //    logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-SampleDataButton_Clicked", "Charts");
    //    $('#btnDrawChartFromSheetData').removeClass('tabButtonActive');
    //    $('#btnDrawChartFromSampleData').addClass('tabButtonActive');
    //    $('#divDrawChart').addClass('activeDrawButton');
    //    $("#startRowTextBox").prop("readonly", false);
    //    $("#endRowTextBox").prop("readonly", false);
    //    $('#divDrawChart').css("color", "#F37A2D");
    //    $('divDrawChart').css("background-color", "white");
    //    hideTooltip();
    //    $('#btnDrawChartFromSheetData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sheet-data-org.png');
    //    $('#btnDrawChartFromSheetData').find('.tabContainer').css('color', 'black');

    //    $('#btnDrawChartFromSampleData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sample-data-active.png');
    //    $('#btnDrawChartFromSampleData').find('.tabContainer').css('color', '#F37A2D');


    //    $('#chkHeaderRow').attr("disabled", "true");
    //    fillDatasourceScreenWithSampleDataDetails();

    //    setDefaultsForDataSourceScreen();
    //    $('.dropdowncolor').css("background-color", "#FDE4D6");
    //    $('.metricdropdowncolor').css("background-color", "#ECEDEF");
    //    //if (selectedChart == "SankeySentimentChart" || selectedChart == "SankeyNonSentimentChart" || selectedChart == "SankeySentimentChartAdvance" || selectedChart == "SankeyNonSentimentChartAdvance") {
    //        insertChartSampleDataIntoSheet();
    //    //}
    //});

    /* Datasource Edit Mode End*/

    //$("#btnDrawChartFromSheetData").on("click", function () {
    //    // alert("syncMode=> "+syncMode +" in btnDrawChartFromSheetData");
    //    if (syncMode == "Edit") {
    //        openSelectedChartGuidInEditMode(synchedChartGUID);
    //        $('#chkHeaderRow').removeAttr("disabled");
    //        return;
    //    }
    //    disableControl();
    //    var chartRuleObject = getSelectedChartDisplayName();
    //    //decide show or hide Dimension/Measure Text
    //    clearTimeout(newchart_added_updated_timer_handler);
    //    logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-SheetDataButton_Clicked", "Charts");
    //    $('#divDrawChart').removeClass('activeDrawButton');
    //    $("#startRowTextBox").prop("readonly", true);
    //    $("#endRowTextBox").prop("readonly", true);
    //    $('#btnDrawChartFromSheetData').addClass('tabButtonActive');
    //    $('#btnDrawChartFromSampleData').removeClass('tabButtonActive');
    //    $('#divDrawChart').css("color", "#B8B8B8");
    //    hideTooltip();

    //    $('#btnDrawChartFromSheetData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sheet-data-active.png');
    //    $('#btnDrawChartFromSheetData').find('.tabContainer').css('color', '#F37A2D');

    //    $('#btnDrawChartFromSampleData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sample-data-org.png');
    //    $('#btnDrawChartFromSampleData').find('.tabContainer').css('color', 'black');

    //    $('#chkHeaderRow').removeAttr("disabled");
    //    $('#chkHeaderRow').prop("checked", "checked");
    //    $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-active.png");
    //    isSampleDataClicked = false;
    //    isEditModeClicked = false;
    //    isEditModeLoadingFromMyChartClick = false;
    //    setDefaultsForDataSourceScreen();
    //    loadDataSourceContainer();
    //});



    $(".maindropDownIcon").click(function (e) {
        $(".settingTab").css("background-color", "#ffffff");
        $(".chartTab").css("background-color", "#f9fbfc");
        $('#propSettingsDivgContiner').hide();
        if ($('#mainDropDownContainer').css("display") == "none") {
            $('#mainDropDownContainer').show();
        }
        else if ($('#mainDropDownContainer').css("display") == "block") {
            $('#mainDropDownContainer').hide();
        }
        if (e) { e.stopPropagation(); }
        if ($('.trialSectionHeaderBar').css('display') == 'none') {
            // $('#mainDropDownContainer').css('top', '15%');
        }
        if ($('.trialSectionHeaderBar').css('display') == 'block') {
            // $('#mainDropDownContainer').css('top', '25%');
        }
    });

    $("#iconMenuHamberger").click(function (e) {
        //alert("open menu clicked");
        if (isMySavedChartExists) {
            $('.mychart').css('cursor', 'pointer');
            $('.mychart').css('color', 'black');
        }
        else {
            $('.mychart').css('cursor', 'default');
            $('.mychart').css('color', '#B8B8B8');
        }
        $(".settingTab").css("background-color", "#ffffff");
        $(".chartTab").css("background-color", "#f9fbfc");
        $('#propSettingsDivgContiner').hide();
        if ($('#mainDropDownContainer').css("display") == "none") {
            $('#mainDropDownContainer').show();
        }
        else if ($('#mainDropDownContainer').css("display") == "block") {
            $('#mainDropDownContainer').hide();
        }
        if (e) { e.stopPropagation(); }
        if ($('.trialSectionHeaderBar').css('display') == 'none') {
            //$('#mainDropDownContainer').css('top', '15%');
        }
        if ($('.trialSectionHeaderBar').css('display') == 'block') {
            //$('#mainDropDownContainer').css('top', '25%');
        }

        logUserActionIntoDatabase("MenuButtonClicked", "Addon");
        $('#tooltipSubscriptionScreen').html('');
        $('#tooltipSubscriptionScreen').hide();
    });

    // temporarily commented - Abid
    //openMyChartsContainerView("ViewMyCharts-LandingPage");
    $('#refreshtxtbox').click(function () {
        $('#myChartsSearchBox').val("");
        openMyChartsContainerView();
    });

    //$(document).on("click", ".clickableExploreSampleData", function (d) {
    //    exploreSampleDataLinkClicked = true;
    //    $('#divDrawChart').trigger("click");
    //});


    $('.ChartListDivs').click(function () {
        $(".selectedChartNameDiv").show();
        //$(".buttonsContainer").show();
    });

    $('#orderMyCharts').click(function () {
        $('#myChartsSearchBox').val("");
        if (currentOrder == "ascending") {
            giveMyChartsOrder("ascending");
            currentOrder = "descending";
            $('#orderMyCharts').attr("title", "Order by ascending");
            $('#orderMyCharts').attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/sort-asc.png");
        }
        else {
            giveMyChartsOrder("descending");
            currentOrder = "ascending";
            $('#orderMyCharts').attr("title", "Order by descending");
            $('#orderMyCharts').attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/sort-desc.png");
        }
    });

    $('#iconMenuBreadcrumb').click(function () {
        screenNavigationHistory.pop();
        var lastScreen = screenNavigationHistory[screenNavigationHistory.length - 1];
        if (lastScreen == "MyChartsScreen") {
            openMyChartsView();
        }
        else if (lastScreen == "ChartListScreen") {
            $('#DataSourceDiv').hide();
            $('.selectedChartNameDiv').hide();
            $("#divChartExpoChartsSearchBox").show();
            $("#divChartExpoCharts").show();
            $('.samplesheetdata').hide();
            $('.rearrangeText').css("z-index", "100");
            $("input").focus();
            if (enableCurrentActiveSheetSelection) {
                //When move from DataSource and click again syncMode changed to Edit instead of Add so manually change it.
                syncMode = "Add";
            }
            $("#licenseKeyContainer").hide();
            $('#priceRangeDetailedScreen').hide();
            $('#editLicenseKeyContainer').hide();
            selectedScreen = "AllChartsScreen";
        }
        else if (lastScreen == "DataSourceScreen") {
            $("#licenseKeyContainer").hide();
            $('.selectedChartNameDiv').show();
            $('#DataSourceDiv').show();
            $('#priceRangeDetailedScreen').hide();
            $('#editLicenseKeyContainer').hide();
            $("#divChartExpoChartsSearchBox").hide();
            $("#divChartExpoCharts").hide();
            selectedScreen = "DataSourceScreen";
        }
        else if (lastScreen == "SubscriptionScreen") {
            var subscriptionTypeVal = $("#radioBtnSingleUserReoccurringMonthlyPayment").attr("subscriptionType");
            if (subscriptionTypeVal == "Monthly") {
                $("#radioBtnSingleUserReoccurringMonthlyPayment").click();
            }
            else {
                $("#radioBtnSingleUserOneMonthPayment").click();
            }
            $('.pageTitle').text('Subscribe');
            $("#licenseKeyContainer").show();
            $('#priceRangeDetailedScreen').hide();
            $('#editLicenseKeyContainer').hide();
            $("#divChartExpoChartsSearchBox").hide();
            $("#divChartExpoCharts").hide();
            selectedScreen = "SubscriptionScreen";
        }
        else if (lastScreen == "PricingScreen") {
            $('.pageTitle').text('ChartExpo Pricing');
            $("#licenseKeyContainer").hide();
            $('#priceRangeDetailedScreen').show();
            $('#editLicenseKeyContainer').hide();
            $("#divChartExpoChartsSearchBox").hide();
            $("#divChartExpoCharts").hide();
            selectedScreen = "PricingScreen";
            var subscriptionTypeVal = $("#autorenewcheckbox").find("#radioBtnSingleUserReoccurringMonthlyPayment").attr("subscriptionType");
            if (subscriptionTypeVal == "Monthly") {
                $("#autorenewcheckbox").find("#radioBtnSingleUserReoccurringMonthlyPayment").click();
            }
            else {
                $("#autorenewcheckbox").find("#radioBtnSingleUserOneMonthPayment").click();
            }
        }
        else if (lastScreen == "TrialUserScreen") {
            $('.pageTitle').text('Manage Trial User(s)');
            $("#licenseKeyContainer").hide();
            $('#priceRangeDetailedScreen').hide();
            $('#editLicenseKeyContainer').show();
            $("#divChartExpoChartsSearchBox").hide();
            $("#divChartExpoCharts").hide();
            selectedScreen = "ManageTeamUsers";
        }
        else if (lastScreen == undefined) {
            openMyChartsView();
        }
        logUserActionIntoDatabase("GoBackButtonClicked", "Addon");
    });

    logUserActionIntoDatabase("LoadingCompleted", "AddonLoading");

    $(".se-pre-con").fadeOut("slow");
    // paste here
});

function initializeDataSourceScreenWithDefaultState() {
    chartRuleObject = getSelectedChartRulesObject();
    selectedchartDisplayName = chartRuleObject.ChartDisplayName;
    //decide show or hide Dimension/Measure Text
    isDimensionMeasureTextVisible(chartRuleObject);
    // TODO Add, category as well along chart name
    logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-Chart_Selected", "Charts");

    $('#btnDrawChartFromSheetData').addClass('tabButtonActive');
    $('#btnDrawChartFromSheetData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sheet-data-active.png');
    $('#btnDrawChartFromSheetData').find('.tabContainer').css('color', '#F37A2D');

    $('#btnDrawChartFromSampleData').removeClass('tabButtonActive');
    $('#btnDrawChartFromSampleData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sample-data-org.png');
    $('#btnDrawChartFromSampleData').find('.tabContainer').css('color', 'black');

    isSampleDataClicked = false;
    isEditModeClicked = false;
    isEditModeLoadingFromMyChartClick = false;

    setDefaultsForDataSourceScreen();
    loadDataSourceContainer();
    viewScreen("DataSourceScreen");
    if (!enableCurrentActiveSheetSelection) {
        disableControl();
    }
    isOriginalTable = false;
    $('#divDrawChart').css('color', '');
}

function resetSearchChartOnNewAction() {
    $('#searchChartExpoCharts').val("");
    var re = new RegExp("", "i"); // "i" means it's case-insensitive
    $('#divChartExpoCharts .chartsList .ChartListDivs').show().filter(function () {
        return !re.test("");
    }).hide();
    $('#divChartExpoCharts .chartsList .ChartListDivsWithScroll').show().filter(function () {
        return !re.test("");
    }).hide();
    searchChartExpoCharts();
}
function collapseCategoryIcons() {
    $('#imgSankey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgSentimentAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgComparativeAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgSpecializedSurvey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgGeneralAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgPPC').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
}

function showSelectChartScreen() {
    resetAllVariables();
    collapseChartsDivSections();
    $('#ChartouterDiv').hide();
    $('#CreateChartContentDiv').hide();
    $('#divChartExpoCharts').show();
    $('#divChartExpoChartsSearchBox').show();
    $("#divSankeyCharts").show();
}

function resetAllVariables() {
    syncMode = "Add";
    chartInEditModeSyncTime = false;
    storableObjectInTempStorage = {
        selectedChart: '',
        selectedChartDisplayName: '',
        editableChartCustomName: '',
        headerRow: '',
        dataRows: '',
        myChart: '',
        editableChartGuid: '',
        dimension: '',
        defaultProperties: '',
        selectedChartCategory: '',
        openDataViewerOnLoad: '',
        sameContractList: '',
        chartAddedUpdatedIntoMyChartList: '',
        headerRowNumber: '',
        headerRowAnnotation: '',
        dataRowsAnnotation: '',
        dataRowFrom: '',
        dataRowTo: '',
        fileName: '',
        fileId: '',
        sheetName: '',
        sheetId: '',
        isSampleData: '',
        sheetWholedataRows: '',
        propertiesSectionHeadsArray: ''
    };
}

function collapseChartsDivSections() {
    $("#divGeneralAnalysisCharts").hide();
    $("#divSpecializedSurveyCharts").hide();
    $("#divComparativeAnalysisCharts").hide();
    $("#divSentimentAnalysisCharts").hide();
    $("#divPPCCharts").hide();
    $('#imgSentimentAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgComparativeAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgSpecializedSurvey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgGeneralAnalysis').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgPPC').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
    $('#imgSankey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
}

function disableControl() {
    $(".positionClass").addClass("disableElemet");
    $(".positionClass").removeClass("enableElemet");
}

function enableControl() {
    $(".positionClass").removeClass("disableElemet");
    $(".positionClass").addClass("enableElemet");
}

function searchChartExpoCharts() {
    var headers = ["#divSankey", "#divSentimentAnalysis", "#divSpecializedSurvey", "#divComparativeAnalysis"];
    var headersWithScrollCharts = ["#divGeneralAnalysis", "#divPPC"];
    for (var i = 0; i < headers.length; i++) {
        $(headers[i]).show();
    }
    for (var i = 0; i < headersWithScrollCharts.length; i++) {
        $(headersWithScrollCharts[i]).show();
    }

    for (var i = 0; i < headers.length; i++) {
        var flag = false;
        var elements = $(headers[i] + "Charts" + " .ChartListDivs");
        for (var j = 0; j < elements.length; j++) {
            if ($(elements[j]).css("display") != "none") {
                flag = true;
                break;
            }
        }
        if (!flag) {
            $(headers[i]).hide();
            $(headers[i] + "Charts").hide();
        }
        else {
            if ($('#searchChartExpoCharts').val() != '') {
                var imgDivText = headers[i].replace('div', 'img');
                $(imgDivText).attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
                $(headers[i] + "Charts").show();
            }
            else {
                collapseCategoryIcons();
                $(headers[i] + "Charts").hide();
                $('#divSankeyCharts').show();
                $('#imgSankey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
            }
        }
    }
    for (var i = 0; i < headersWithScrollCharts.length; i++) {
        var flag = false;
        var elements = $(headersWithScrollCharts[i] + "Charts" + " .ChartListDivsWithScroll");
        for (var j = 0; j < elements.length; j++) {
            if ($(elements[j]).css("display") != "none") {
                flag = true;
                break;
            }
        }
        if (!flag) {
            $(headersWithScrollCharts[i]).hide();
            $(headersWithScrollCharts[i] + "Charts").hide();
        }
        else {
            if ($('#searchChartExpoCharts').val() != '') {
                var imgDivText = headersWithScrollCharts[i].replace('div', 'img');
                $(imgDivText).attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
                $(headersWithScrollCharts[i] + "Charts").show();
            }
            else {
                collapseCategoryIcons();
                $(headersWithScrollCharts[i] + "Charts").hide();
                $('#divSankeyCharts').show();
                $('#imgSankey').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropup-icon.png');
            }
        }
    }
}

function openHelpInNewTab() {
    //window.open('http://chartexpo.com/home/googlesheetaddonhelp', '_blank');
    window.open('https://chartexpo.com/home/addonhelp#googlesheet', '_blank');
}

/**
 * Datasource start
 */
//var dataSourceColumns = [["Topic", "Subtopic-1", "Subtopic-2", "Subtopic-3", "Subtopic-4", "Subtopic-5", "Count"]];

var headerColumnsWithEmptyName = [];

var headerRowNumber = 0;

//var isSampleDataClicked = false;
//var isEditModeClicked = false;
//var isEditModeLoadingFromMyChartClick = false;
var isSearchMenuOpenedFromDimensionAddButton = false;
var isSearchMenuOpenedFromMeasureAddButton = false;

function onHeaderRowCheckboxChange(obj) {
    if (obj.checked) {
        $("#txtBoxHeaderRow").attr("min", "1");
        $("#txtBoxHeaderRow").val("1");
        $("#txtBoxHeaderRow").removeAttr("disabled");
    }
    else {
        $("#txtBoxHeaderRow").attr("min", "0");
        $("#txtBoxHeaderRow").val("0");
        $("#txtBoxHeaderRow").attr("disabled", "disabled");
    }
}

function addDimensionsDropDownList(columnSelected, isSettingsColumnExistInSheet) {
    if (dataSourceColumns.length > 0) {

        var deletedColumnColor = "";
        var makeColumnDraggable = 'draggable="true"';
        if (isSettingsColumnExistInSheet != null && isSettingsColumnExistInSheet == false) {

            deletedColumnColor = "background-color:#f08080;";
            makeColumnDraggable = "";
        }
        var threedigitsrandom = Math.floor(100 + Math.random() * 900);
        var dimensionsInputLength = $('.dimension').length;
        var allowedDimensions = getSelectedChartAllowedDimensions();

        var webkitAprearance = '';
        if (isSampleDataClicked) {
            webkitAprearance = '-webkit-appearance:none';
        }

        var leftSpace = ' style=margin-bottom:4px;'
        var isShowLeftMargin = dimensionsInputLength % 2;
        if (isShowLeftMargin != 0) {
            leftSpace += 'margin-left:7px;'
            $('#addDimensionClickContainer').css('margin-left', '0px');
        }
        else {
            $('#addDimensionClickContainer').css('margin-left', '7px');
        }
        var dimensionsDropdownList = '<div column="dimension" class="dimension" ' + makeColumnDraggable + leftSpace + '  > <div id="dropDownDimensions' + dimensionsInputLength + threedigitsrandom + '" class="DropdownList dropdowncolor dropdownchangeclass" style="width:100px;padding-left:4px;float:left;' + deletedColumnColor + '" isDeletedColumn=' + (isSettingsColumnExistInSheet == false ? "true" : "false") + '>';

        var firstAddedColumn = columnSelected;
        dimensionsDropdownList = dimensionsDropdownList + '<div style="padding-top:4px;float:left;width: 83px;white-space: nowrap;overflow: hidden;text-overflow:ellipsis;" value="' + columnSelected + '">' + columnSelected + '</div >';
        dimensionsDropdownList = dimensionsDropdownList + '<div style="float:right;padding-top:4px;padding-right:5px;"><img src="https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png"></div>';


        if (firstAddedColumn == '') {
            //alert('No more dimension/measure exist!');
            showMessageDialog("Dimension/Measure", "No more dimension/measure exist!", "confirmation", false, [], true);
            return;
        }
        //Updating selectedChartColumns list.
        if (firstAddedColumn != '') {
            selectedChartColumns.push(firstAddedColumn);
            //get A1Notation information.
            var columnA1NotationInformation = getColumnA1NotationFromSheetData(firstAddedColumn);
            if (columnA1NotationInformation != "") {
                //selectedChartA1NotationColumns.push(columnA1NotationInformation);
                updateA1NotationColumn("", columnA1NotationInformation);
                //highlightSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumns);
                updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);

            }
        }

        dimensionsDropdownList = dimensionsDropdownList + '</div><img id="dimensionRemove' + dimensionsInputLength + '" style="margin-left:5px;padding-top:6px;" class="imagecontainer dimensionRemoveClass" src="https://apps.polyvista.com/GooglesheetFeb2021/Scripts/Polyvista/feedback177/icons/remove-selection.svg" title="remove" /></div>';

        $('#addDimensionClickContainer').before($(dimensionsDropdownList));
        attachClickEvent('dimensionRemoveClass', 'dimension');
        attachChangeEvent('dropdownchangeclass');

        dimensionsInputLength = $('.dimension').length;
        if (dimensionsInputLength == allowedDimensions) {
            $('#addDimensionClick').css('background', 'white');
        }

        $(".dimension").unbind("click");
        $(".dimension").on('click', function () {
            showSearchMenuOnDropdownClick(this, event);
        });
        $(".dropdownchangeclass").on('focus', function () {
            $(this).attr("disabled", "true");
            showSearchMenuOnDropdownClick(this.parentElement);
        });

        //set row slider to maximum value for the first time when dimension or measure is added.
        if (dimensionsInputLength == 1 && $('.metric').length == 0) {
            updateSliderValue(0, sheetRecords.length);
        }
    }
}

function bindDivsDropEvent() {
    var dragged;

    /* events fired on the draggable target */
    document.addEventListener("drag", function (event) {

    }, false);

    document.addEventListener("dragstart", function (event) {
        // store a ref. on the dragged elem
        dragged = event.target;
        // make it half transparent
        //event.target.style.opacity = .5;
    }, false);

    document.addEventListener("dragend", function (event) {
        // reset the transparency
        //event.target.style.opacity = "";
    }, false);

    /* events fired on the drop targets */
    document.addEventListener("dragover", function (event) {
        // prevent default to allow drop
        event.preventDefault();
    }, false);

    document.addEventListener("dragenter", function (event) {
        // highlight potential drop target when the draggable element enters it
        if (event.target.className == "dimension") {
            //event.target.style.background = "purple";
        }

    }, false);

    document.addEventListener("dragleave", function (event) {
        // reset background of potential drop target when the draggable element leaves it
        if (event.target.className == "dimension") {
            event.target.style.background = "";
        }

    }, false);

    //removeDropEvent('dimension');
    attachDropEvent('dimension');
    attachDropEvent('metric');


    function dropEventBinding() {
        //move dragged elem to the selected drop target
        //alert(event.target.className);

        var isTargetColumnDeleted = $(event.target).attr('isdeletedcolumn');
        if (dragged.children[0].className == event.target.className &&
            (event.target.className == "DropdownList dropdowncolor dropdownchangeclass" || event.target.className == "DropdownList metricdropdowncolor dropdownchangeclass")
            && isTargetColumnDeleted != "true" &&
            $(dragged.children[0]).attr('id') != $(event.target).attr('id')) {
            var targetNode = event.target.parentNode;

            $(dragged).insertBefore(targetNode);
        }
        else if (dragged.children[0].className == $(event.target).parent().attr('class') &&
            ($(event.target).parent().attr('class') == "DropdownList dropdowncolor dropdownchangeclass" || $(event.target).parent().attr('class') == "DropdownList metricdropdowncolor dropdownchangeclass")
            && $(event.target).parent().attr('isdeletedcolumn') != "true" &&
            $(dragged.children[0]).attr('id') != $(event.target).parent().attr('id')) {

            var targetNode = event.target.parentNode.parentNode;

            $(dragged).insertBefore(targetNode);
        }

        setLevelsAndMetricsPosition();
        event.stopImmediatePropagation();
    }

    function setLevelsAndMetricsPosition() {
        $(".dimension").removeAttr("style");
        $(".dimension").each(function (index) {
            if (index % 2 != 0) {
                $(this).css("margin-bottom", "4px");
                $(this).css("margin-left", "7px");
            }
            else {
                $(this).css("margin-bottom", "4px");
            }
        }); //end of each loop.
        $(".metric").removeAttr("style");
        $(".metric").each(function (index) {
            if (index % 2 != 0) {
                $(this).css("margin-bottom", "4px");
                $(this).css("margin-left", "7px");
            }
            else {
                $(this).css("margin-bottom", "4px");
            }
        }); //end of each loop.
    }

    function removeDropEvent(className) {
        var divList;
        // get all the elements with className 'btn'. It returns an array
        var divList = document.getElementsByClassName(className);
        // get the lenght of array defined above
        var listLength = divList.length;
        var i = 0;
        // run the for look for each element in the array
        for (; i < listLength; i++) {
            // attach the event listener                  
            divList[i].removeEventListener("drop", dropEventBinding);
        }
    }
    function attachDropEvent(className) {
        var divList;
        // get all the elements with className 'btn'. It returns an array
        var divList = document.getElementsByClassName(className);
        // get the lenght of array defined above
        var listLength = divList.length;
        var i = 0;
        // run the for look for each element in the array
        for (; i < listLength; i++) {
            // attach the event listener                  
            divList[i].addEventListener("drop", dropEventBinding);
        }
    }

}

function getHeaderNewColumns(clickedObject) {
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    var sheetName = $('#dropdownSheets').val();
    var headerRowNumber = $("#txtBoxHeaderRow").val();
    //console.log('get header new columns' + headerRowNumber);
    if (clickedObject != null) {
        showColumnSearchMenu(clickedObject);
    }
    $(".se-pre-con-loader").fadeIn("slow");
    $("#chartexpo_GoogleSheetAddon_tileMenu_txtSearch").focus();
    //Get header columns of selected sheet.
    if (headerRowNumber > 0) {
        google.script.run.withSuccessHandler(function (columns) {
            //console.log("header columns:" + headerColumns);
            columns = JSON.parse(columns);
            headerColumns = columns.columnNames;
            headerColumnsWithEmptyName = columns.columnNamesWithEmptyName;
            if (headerColumns != null && headerColumns != undefined && headerColumns.length > 0) {
                MapOldHeaderColumnsWithNewHeaderColumns(headerColumns);
                loadOriginalColumns(clickedObject);
            }
            $(".se-pre-con-loader").fadeOut("slow");
            $('.dropdownchangeclass').removeAttr('disabled');

        })
            .withFailureHandler(
            function (msg, element) {
                $('.dropdownchangeclass').removeAttr('disabled');
                $(".se-pre-con-loader").fadeOut("slow");
                handleError(msg);
            }).getSelectiveHeaderRowDetail(sheetName, headerRowNumber);
    }
    else {
        loadOriginalColumns(clickedObject);
        $(".se-pre-con-loader").fadeOut("slow");
        $('.dropdownchangeclass').removeAttr('disabled');
    }
}

function MapOldHeaderColumnsWithNewHeaderColumns(newHeaderColumns) {
    //debugger;
    var processedElement = [];
    //newHeaderColumns = [{ColumnName: "topic", A1NotationDetail: "A1", Row:3, Col:1 },
    //    {ColumnName: "subtopic", A1NotationDetail: "B1", Row: 3, Col: 2 }, {ColumnName: "subtopic1", A1NotationDetail: "C1", Row: 3, Col: 3},
    //    {ColumnName: "subtopic", A1NotationDetail: "D1", Row: 3, Col: 4 }, {ColumnName: "subtopic7", A1NotationDetail: "E1", Row: 3, Col: 5 },
    //    {ColumnName: "subtopic1", A1NotationDetail: "F1", Row: 3, Col: 6},
    //    {ColumnName: "count", A1NotationDetail: "G1", Row: 3, Col: 7}
    //];
    //datasourceColumnsWithIndex = [{SheetColumnIndex: 0, ColumnName: "topic", A1NotationDetail: "A1" },
    //    {SheetColumnIndex: 1, ColumnName: "subtopic", A1NotationDetail: "B1" }, {SheetColumnIndex: 2, ColumnName: "subtopic1", A1NotationDetail: "C1" },
    //    {SheetColumnIndex: 3, ColumnName: "subtopic2", A1NotationDetail: "D1" }, {SheetColumnIndex: 4, ColumnName: "count", A1NotationDetail: "E1" },
    //    {SheetColumnIndex: 5, ColumnName: "count", A1NotationDetail: "F1" }  ];

    if (newHeaderColumns.length == datasourceColumnsWithIndex.length) {
        //if length is same then check whether existing columns are modified.
        var isColumnsNameMatched = matchNewColumnsWithOldColumns(newHeaderColumns);
        //if columns are matched then there is no change in sheet columns and there is no need to update object structure.

        if (!isColumnsNameMatched) {
            processNewHeaderColumns(newHeaderColumns, processedElement, "true");
        }
        else if ($('.dimension').text() == "" && $('.metric').text() == "") {
            processNewHeaderColumns(newHeaderColumns, processedElement);
        }
        else {
            if (isEditModeClicked) {
                processNewHeaderColumns(newHeaderColumns, processedElement, "true");
            }
        }
    }
    else { //If columns length is not same then its mean columns are added.
        processNewHeaderColumns(newHeaderColumns, processedElement, "true");
    }
}

function processNewHeaderColumns(newHeaderColumns, processedElement, notResetObjectStructure) {
    //Update object structure with new and changed columns and reset datasource screen.
    var duplicateColumns = [];
    if (notResetObjectStructure != "true") {
        resetObjectStructrueAfterHeaderColumnsChange();
    }
    //Also set dataSourceColumns[0], sheetA1NotationDetails and datasourceColumnsWithIndex with new header columns.
    //sheetA1NotationDetails object is updated on create chart button click.
    if (newHeaderColumns != null && newHeaderColumns.length > 0) {
        duplicateColumns = getDuplicateColumns(newHeaderColumns, processedElement);
    }
    //remove duplicate column names from newHeaderColumns.
    //processedElement array contains duplicate columns.
    if (processedElement.length > 0) {
        //removing duplicate columns.
        newHeaderColumns = removeDuplicateColumnsFromNewHeaderColumns(newHeaderColumns, processedElement);
    }
    //Add duplicate columns in newHeaderColumns.

    if (duplicateColumns.length > 0) {
        //Add duplicate columns.
        newHeaderColumns = newHeaderColumns.concat(duplicateColumns);
        dataSourceDuplicateColumns = duplicateColumns;
        //alert(JSON.stringify(dataSourceDuplicateColumns));
    }

    //Filling datasource object with new and updated header columns.
    if (newHeaderColumns != null && newHeaderColumns.length > 0) {
        fillDatasourceObjectsWithNewHeaderColumns(newHeaderColumns);
    }
}

//This is used to remove duplicate columns from newHeaderColumns.
function removeDuplicateColumnsFromNewHeaderColumns(newHeaderColumns, processedElement) {
    for (var counter = 0; counter < newHeaderColumns.length; counter++) {
        for (var processedElementIndex = 0; processedElementIndex < processedElement.length; processedElementIndex++) {
            if (newHeaderColumns[counter].ColumnName == processedElement[processedElementIndex]) {
                newHeaderColumns.splice(counter, 1);
                counter = counter - 1;
                break;
            }
        } //end of inner for loop.

    }//end of outer for loop.
    return newHeaderColumns;
}

//matching new columns with old columns before change on sheet.
function matchNewColumnsWithOldColumns(newHeaderColumns) {
    var isColumnMatched = false;
    for (var counter = 0; counter < newHeaderColumns.length; counter++) {
        isColumnMatched = false;

        for (var datasourceColumnIndex = 0; datasourceColumnIndex < datasourceColumnsWithIndex.length; datasourceColumnIndex++) {
            if (newHeaderColumns[counter].ColumnName == datasourceColumnsWithIndex[datasourceColumnIndex].ColumnName) {
                isColumnMatched = true;
                break;
            }
        } //end of for loop.

        //If column is not matched in datasourceColumnsWithIndex the check it in duplicateColumns.
        if (!isColumnMatched) {
            var columnName = newHeaderColumns[counter].ColumnName + "_" + newHeaderColumns[counter].A1NotationDetail;
            for (var duplicateColumnIndex = 0; duplicateColumnIndex < dataSourceDuplicateColumns.length; duplicateColumnIndex++) {

                if (columnName == dataSourceDuplicateColumns[duplicateColumnIndex].ColumnName) {
                    isColumnMatched = true;
                    break;
                }
            }// end of for loop.
        } //end of if statement.

        //if column is not matched.
        if (!isColumnMatched) {
            return false;
        }
    }//end of for loop.

    return true;
}

//This method is getting duplicate columns.
function getDuplicateColumns(columns, processedElement) {
    var columnsCopy = JSON.parse(JSON.stringify(columns));
    var duplicateColumns = [];
    //var processedElement = [];
    for (var counter = 0; counter < columns.length; counter++) {
        var isProcessed = checkArrayElementProcessed(processedElement, columns[counter].ColumnName);
        if (!isProcessed) {//If array element is not processed.
            var duplicates = checkDuplicateColumns(columns[counter], columnsCopy);
            if (duplicates.length > 0) {
                //If duplicate is found then remove duplicate elements from columnsCopy array.                   
                var filtered = columnsCopy.filter(function (element, index, arr) { return element.ColumnName != columns[counter].ColumnName; });
                columnsCopy = filtered;
                //Add processed column to avoid proocessing of same column name next time in loop.
                processedElement.push(columns[counter].ColumnName);
                //concatenate duplicate columns.
                duplicateColumns = duplicateColumns.concat(duplicates);
            }
        }
    }//end of for loop.

    return duplicateColumns;
}

function checkArrayElementProcessed(processedElement, columnName) {
    for (var index = 0; index < processedElement.length; index++) {
        if (processedElement[index] == columnName) {
            return true;
        }
    }
    return false;
}

function checkDuplicateColumns(columnObject, columnsCopy) {
    var duplicates = [];
    var duplicateNumbers = 0;
    for (var counter = 0; counter < columnsCopy.length; counter++) {
        if (columnsCopy[counter].ColumnName == columnObject.ColumnName) {
            duplicateNumbers = duplicateNumbers + 1;
        }

        if (columnsCopy[counter].ColumnName == columnObject.ColumnName && duplicateNumbers > 1) // if duplicate found.
        {
            if (duplicateNumbers == 2) {
                duplicates.push({
                    ColumnName: columnObject.ColumnName + "_" + columnObject.A1NotationDetail, A1NotationDetail: columnObject.A1NotationDetail, OldColumnName: columnObject.ColumnName,
                    Col: columnObject.Col, Row: columnObject.Row, A1NotationInRCFormat: columnObject.A1NotationInRCFormat
                });
                duplicates.push({
                    ColumnName: columnsCopy[counter].ColumnName + "_" + columnsCopy[counter].A1NotationDetail, A1NotationDetail: columnsCopy[counter].A1NotationDetail,
                    OldColumnName: columnsCopy[counter].ColumnName, Col: columnsCopy[counter].Col, Row: columnsCopy[counter].Row, A1NotationInRCFormat: columnsCopy[counter].A1NotationInRCFormat
                });
            }
            else {
                duplicates.push({
                    ColumnName: columnsCopy[counter].ColumnName + "_" + columnsCopy[counter].A1NotationDetail, A1NotationDetail: columnsCopy[counter].A1NotationDetail,
                    OldColumnName: columnsCopy[counter].ColumnName, Col: columnsCopy[counter].Col, Row: columnsCopy[counter].Row, A1NotationInRCFormat: columnsCopy[counter].A1NotationInRCFormat
                });
            }
        }

    } //end of for loop.
    return duplicates;
}

function resetObjectStructrueAfterHeaderColumnsChange() {
    dataSourceColumns[0] = [];
    dataSourceColumnsAll[0] = []
    sheetA1NotationDetails = [];
    datasourceColumnsWithIndex = [];
    selectedChartColumns = [];
    selectedChartA1NotationColumns = [];

    dataSourceDuplicateColumns = [];
    //resetting selected dimension and measures on header columns change.
    $('.dimension').remove();
    $('.metric').remove();
}

function fillDatasourceObjectsWithNewHeaderColumns(newHeaderColumns) {
    var filteredColumns = [];
    var allColumns = [];
    datasourceColumnsWithIndex = [];
    for (var counter = 0; counter < newHeaderColumns.length; counter++) {
        allColumns.push(newHeaderColumns[counter].ColumnName);
        if (newHeaderColumns[counter].ColumnName != null && newHeaderColumns[counter].ColumnName != "") {
            filteredColumns.push(newHeaderColumns[counter].ColumnName);
            datasourceColumnsWithIndex.push({
                SheetColumnIndex: newHeaderColumns[counter].Col - 1,
                ColumnName: newHeaderColumns[counter].ColumnName, A1NotationDetail: newHeaderColumns[counter].A1NotationDetail,
                A1NotationInRCFormat: newHeaderColumns[counter].A1NotationInRCFormat
            });
        }
    } //end of for loop.

    dataSourceColumns[0] = filteredColumns;//setting dataSourceColumns object.
    dataSourceColumnsAll[0] = allColumns;
}

function isColumnExistInDuplicateColumns(columnName) {
    for (var counter = 0; counter < dataSourceDuplicateColumns.length; counter++) {
        if (dataSourceDuplicateColumns[counter].ColumnName == columnName) {
            return true;
        }
    }//end of for loop.
    return false;
}

var searchMenuPreviousSelectedValue;
var searchMenuSelectedObject;
function showSearchMenuOnDropdownClick(obj, event) {
    isSearchMenuOpenedFromDimensionAddButton = false;
    isSearchMenuOpenedFromMeasureAddButton = false;
    searchMenuPreviousSelectedValue = $(obj).find('div').first().text();
    searchMenuSelectedObject = $(obj).find('div').find('div').first();
    //alert(searchMenuPreviousSelectedValue);

    $('.dropdownchangeclass').attr("disabled", "true");


    var isColumnDeleted = $(obj).find('div').first().attr("isdeletedcolumn");
    //If column is deleted then also don't show search menu.

    if (!isSampleDataClicked) {//If it is with sample data then don't show search menu and also for deleted columns.
        //loadOriginalColumns();
        //showColumnSearchMenu(obj);
        getHeaderNewColumns(obj);
        //$('.dropdownchangeclass').removeAttr('disabled');
    }
    event.stopPropagation();
}

function showColumnSearchMenu(clickedObject) {
    $("#uloriginalcolumns").empty();
    //$("#uloriginalcolumns").append($('<div>Original Columns</div>'));
    $("#ulduplicatecolumns").empty();

    $('#chartexpo_GoogleSheetAddon_tileMenu_txtSearch').val('');
    var leftPositionToMove = 40;
    var topPosition = $(clickedObject).position().top + $('#DataSourceDiv').scrollTop();
    var leftPosition = $(clickedObject).position().left;
    $("#myDropdown").css({ top: topPosition + 25, left: leftPosition < 100 ? leftPosition : leftPosition - leftPositionToMove });
    document.getElementById("myDropdown").classList.toggle("show");
    selectedDimensionMeasure = clickedObject;
    if (event != undefined) {
        event.stopPropagation();
    }
}

function showSheetsSearchMenu(clickedObject) {

    $('#chartexpo_GoogleSheetAddon_tileMenu_txtSearch_selectSheet').val('');
    var leftPositionToMove = 13;
    var topPosition = $(clickedObject).position().top + $('#DataSourceDiv').scrollTop();
    var leftPosition = $(clickedObject).position().left;
    $("#myDropdownSelectSheet").css({ top: topPosition + 28, left: leftPosition + leftPositionToMove });
    if (event != undefined) {
        event.stopPropagation();
    }
}


function hideSearchMenu() {
    $("#myDropdown").removeClass("show");
}
function convertChartDataColumnsIntoDraggableFormat(draggableColumnsArray, tblDefaultHeader, tblDefaultBody) {
    var draggableBodyRows = [];
    var headerColumns = tblDefaultHeader;
    var draggableColumns = draggableColumnsArray;
    for (var i = 0; i < draggableColumns[0].length; i++)//Subtopic-2,Topic,Subtopic-1,Count
    {
        for (var j = 0; j < headerColumns[0].length; j++)//Topic,Subtopic-1,Subtopic-2,Count
        {
            if (draggableColumns[0][i] == headerColumns[0][j]) {
                var draggableSingleRow = [];
                for (var k = 0; k < tblDefaultBody.length; k++) {
                    draggableSingleRow.push(tblDefaultBody[k][j]);
                    //tblDefaultBody[k][j] = temp;
                }
                draggableBodyRows.push(draggableSingleRow)
            }
        }
    }

    return convertColumnArraysIntoTabularForm(draggableBodyRows);
}
function convertColumnArraysIntoTabularForm(draggableBodyRows) {
    var formatedBodyRows = [];
    for (var i = 0; i < draggableBodyRows[0].length; i++)//Subtopic-2,Topic,Subtopic-1,Count
    {
        var localArray = [];
        for (var j = 0; j < draggableBodyRows.length; j++)//Topic,Subtopic-1,Subtopic-2,Count
        {
            localArray.push(draggableBodyRows[j][i])
        }
        formatedBodyRows.push(localArray);
    }
    return formatedBodyRows;
}
function insertChartSampleDataIntoSheet(openedFromSource) {
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    var sheetNumber = 1;
    logUserActionIntoDatabase(selectedChart + "-SampleSheetAdded_" + openedFromSource, "Charts");
    var sheetName = SampleChartCreationSteps[selectedChart]();
    var timeStamp = new Date();
    var chartTitle = [[sheetName[0].chartName]];// + " SampleData"
    var headerColumns = [columnNameMapper[selectedChart]];
    var headerRows = [];
    var stepsHeader = [[sheetName[0].sheetHeaderTitle]];
    var stepsContent = sheetName[0].steps;
    var videoLink = [["https://www.youtube.com/watch?v=Do0Mp88hszQ"]];
    var imagePath = "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/chartsampleimage/" + selectedChart + ".png";

    if (selectedChartCategory !== undefined && selectedChartCategory === "PPC") {
        headerRows = PPCChartsSampleData[selectedChart]();
        //If catagory is PPC then these seven repeated charts path is with -PPC
        if (selectedChart == "BarChart" || selectedChart == "ColumnChart" || selectedChart == "GroupedBarChart" || selectedChart == "GroupedColumnChart" || selectedChart == "RadarChart" || selectedChart == "DonutChart" || selectedChart == "BarStackedComparisonChart") {
            imagePath = "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/chartsampleimage/" + selectedChart + "-PPC.png";
        }
    }
    else {
        headerRows = SampleData[selectedChart]();
    }
    if (selectedChart == "ParetoGroupedChart" || selectedChart == "ParetoGroupedHorizontalChart") {
        imagePath = "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/chartsampleimage/" + selectedChart + ".png";
    }
    var tblHeader = [];
    var tblBody = [];
    if (selectedChart == "SentimentTrendChart") {
        tblBody = sentimentTrendGetDataRowForExcel(headerRows);
    }
    else if (selectedChart == "ParetoGroupedChart" || selectedChart == "ParetoGroupedHorizontalChart") {
        tblBody = paretoGroupedChartGetDataRowForExcel(headerRows);
    }
    else if (tabularDataCharts_DimensionColumnName.hasOwnProperty(selectedChart)) {
        tblBody = convertDataIntoExcelFormat(headerRows);
    }
    // for multi measure charts call MultiMeasure method
    else if (selectedChart == "GaugeChart" || selectedChart == "ISGraph" || selectedChart == "ScatterChartAdvance" ||
        selectedChart == "HierarichalBarChartAdvance" || selectedChart == "ChordChart" || selectedChart == "DoubleMeasureComparisonChart") {
        if (selectedChart == "GaugeChart") {
            var gaugeData = [];
            gaugeData.push(headerRows)
            headerRows = gaugeData;
        }
        tblBody = getDataRowForExcelMultiMeasure(headerRows);
    }
    else {
        tblBody = getDataRowForExcel(headerRows);
    }
    if (tabularDataCharts_DimensionColumnName.hasOwnProperty(selectedChart)) {
        tblHeader = [tblBody[0]];
        tblBody = tblBody.slice(1)
    }
    //convert data object into 2d array
    //convert data object into 2d array
    else {
        for (var i = 0; i < headerColumns.length; i++) {
            tblHeader.push(headerColumns[i]);
        }
    }
    var maxValue = 0;
    for (var i = 0; i < sheetNames.length; i++) {
        if (sheetNames[i].match(sheetName[0].chartName + "Sample")) {
            var sheetCurrentValue = sheetNames[i].replace(sheetName[0].chartName + "Sample", "");
            if (maxValue < parseInt(sheetCurrentValue)) {
                maxValue = parseInt(sheetCurrentValue);
            }
        }
    }
    sheetNumber = maxValue + 1;
    if (sheetNumber < 10) {
        sheetNumber = "0" + sheetNumber;
    }
    $(".se-pre-con").fadeIn("slow");
    sheetNames.push(sheetName[0].chartName + "Sample" + sheetNumber);
    google.script.run.withSuccessHandler(function (newlyInsertedSampleSheetDetail) {
        //alert("after sample sheet insertion => " + newlyInsertedSampleSheetDetail);
        //$(".se-pre-con").fadeOut("slow");

        // add new charts into meta and then set it default
        if (newlyInsertedSampleSheetDetail != null && newlyInsertedSampleSheetDetail.length > 0) {

            tblHeaderOfSampleSheet = tblHeader;
            tblBodyOfSampleSheet = tblBody;
            newlyInsertedSampleSheetCompleteDetail = JSON.parse(newlyInsertedSampleSheetDetail);

            // addSampleSheetChartIntoMyChartList(tblHeader, tblBody, JSON.parse(newlyInsertedSampleSheetDetail));
            //if (selectedChart=="")
            if (selectedChart == "SankeySentimentChartAdvance" || selectedChart == "SankeyNonSentimentChartAdvance") {
                addSampleSheetChartIntoMyChartList();
            }
            else {
                //AddScript("https://chartexpo.com/ChartExpoForGoogleSheetAddin/Scripts/Polyvista/ChartExpo.February.v1221.js", "ChartExpo", addSampleSheetChartIntoMyChartList, true);//runInsertChartImageIntoSheetMethod, true);
                AddScript("https://doc-04-ao-docs.googleusercontent.com/docs/securesc/4rk5cof9mh9g2noe3od806r4a21a2tb9/3afilur1j7n84qh6uvjqkfo0ceiq57qv/1619440800000/10040362025876787199/10040362025876787199/1EklpkOGdEzqX8jn1fnUQncQwVBH_HOmQ?e=download&authuser=0&nonce=0ru7594u98hjm&user=10040362025876787199&hash=4p9pg90kqf5h9fpoocokrgomep4cqim6", "ChartExpo", addSampleSheetChartIntoMyChartList, true);//runInsertChartImageIntoSheetMethod, true);
            }
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        })
        .addSampleTestSheet(sheetName[0].chartName + "Sample" + sheetNumber, JSON.stringify(chartTitle), JSON.stringify(tblHeader), JSON.stringify(tblBody), JSON.stringify(stepsHeader), JSON.stringify(stepsContent), imagePath, sheetName[0].chartName, JSON.stringify(videoLink));
    $('#dropdownSheets').append($('<option value="' + sheetName[0].chartName + "Sample" + sheetNumber + '">' + sheetName[0].chartName + "Sample" + sheetNumber + '</option>'));

}

// these variables contains values of data of sheet data
var tblHeaderOfSampleSheet, tblBodyOfSampleSheet, newlyInsertedSampleSheetCompleteDetail;

function loadOriginalColumns(clickedObject) {
    var isDimensionCicked = null;
    if (clickedObject != null && clickedObject != undefined) {
        isDimensionCicked = $(clickedObject).attr("column");
    }
    //$("#uloriginalcolumns").empty();
    //$("#uloriginalcolumns").append($('<div>Original Columns</div>'));
    $("#ulColumnsContainer").empty();
    if (isDimensionCicked == "metric") {
        $("#ulColumnsContainer").append($('<ul id="ulmetriccolumns"></ul>'));
        $("#ulColumnsContainer").append($('<ul id="uldimensioncolumns"></ul>'));
    }
    if (isDimensionCicked == "dimension") {
        $("#ulColumnsContainer").append($('<ul id="uldimensioncolumns"></ul>'));
        $("#ulColumnsContainer").append($('<ul id="ulmetriccolumns"></ul>'));
    }
    //$("#ulduplicatecolumns").empty();
    if (dataSourceDuplicateColumns.length > 0) {
        updateDataSourceColumnsAll();
    }
    if (headerColumnsWithEmptyName.length > dataSourceColumnsAll[0].length) {
        var indexArray = headerColumnsWithEmptyName.filter(function (d, i) {
            if (d.ColumnName == "") {
                return d;
            }
        });
        for (var i = 0; i < indexArray.length; i++) {
            dataSourceColumnsAll[0].splice(indexArray[i].Col - 1, 0, "");
        }
    }
    var startRowIndex = +($("#startRowTextBox").val());
    var numberOfRows = 100;
    //console.log("actualSheetData", JSON.stringify(actualSheetData));
    var selectedData_From_ActualSheetData = actualSheetData.slice(startRowIndex, startRowIndex + numberOfRows);
    if (startRowIndex == actualSheetData.length) {
        selectedData_From_ActualSheetData = [actualSheetData[startRowIndex - 1]];
    }
    //console.log("startRowIndex", startRowIndex);
    //console.log("actualSheetData.length", actualSheetData.length);
    //console.log("selectedData_From_ActualSheetData", JSON.stringify(selectedData_From_ActualSheetData));
    //console.log("dataSourceColumnsAll[0]", JSON.stringify(dataSourceColumnsAll[0]));
    //console.log("dataSourceColumns[0]", JSON.stringify(dataSourceColumns[0]));
    //console.log("dataSourceDuplicateColumns", JSON.stringify(dataSourceDuplicateColumns));

    dimensionColumns = getDimensionColumns(selectedData_From_ActualSheetData);//['Col_1','Col_2','Col_3','Col_4','Col_5','Col_6','Col_7','Col_8'];
    metricColumns = getMetricColumns(selectedData_From_ActualSheetData);//['Metric_9','Metric_10'];

    if (dataSourceColumns.length > 0) {
        for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
            if (!isColumnExistInSelectedCollection(dataSourceColumns[0][counter]))//If column does not exist in selectedChartColumns collection then add it.
            {
                //Also check column should not be in dataSourceDuplicateColumns collection then add it.
                if (!isColumnExistInDuplicateColumns(dataSourceColumns[0][counter])) {
                    if (metricColumns.indexOf(dataSourceColumns[0][counter]) == -1) {
                        var totalColumns = $("#uldimensioncolumns").find('div');
                        if (totalColumns.length == 0) {
                            $("#uldimensioncolumns").append($('<div><b>Dimensions</b></div>'));
                        }
                        $("#uldimensioncolumns").append($('<li class="chartexpo_GoogleSheetAddon_tileMenu_li" originaltext="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</li>'));
                    }
                    if (dimensionColumns.indexOf(dataSourceColumns[0][counter]) == -1) {
                        var totalColumns = $("#ulmetriccolumns").find($('div'));
                        if (totalColumns.length == 0) {
                            $("#ulmetriccolumns").append($('<div><b>Metrics</b></div>'));
                        }
                        $("#ulmetriccolumns").append($('<li class="chartexpo_GoogleSheetAddon_tileMenu_li" originaltext="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</li>'));
                    }
                }//end of inner if statement.
            }//end of outer if statement.
        }//end of for loop.
        //alert('duplicate': dataSourceDuplicateColumns.length);
        //Adding duplicate columns in search menu control.
        if (dataSourceDuplicateColumns.length > 0) {
            for (var counter = 0; counter < dataSourceDuplicateColumns.length; counter++) {
                if (!isColumnExistInSelectedCollection(dataSourceDuplicateColumns[counter].ColumnName)) {
                    if (metricColumns.indexOf(dataSourceDuplicateColumns[counter].ColumnName) == -1) {
                        var totalColumns = $("#uldimensioncolumns").find($('div'));
                        if (totalColumns.length == 0) {
                            $("#uldimensioncolumns").append($('<div><b>Dimensions</b></div>'));
                        }
                        $("#uldimensioncolumns").append($('<li class="chartexpo_GoogleSheetAddon_tileMenu_li" originaltext="' + dataSourceDuplicateColumns[counter].ColumnName + '">' + dataSourceDuplicateColumns[counter].ColumnName + '</li>'));
                    }
                    if (dimensionColumns.indexOf(dataSourceDuplicateColumns[counter].ColumnName) == -1) {
                        var totalColumns = $("#ulmetriccolumns").find($('div'));
                        if (totalColumns.length == 0) {
                            $("#ulmetriccolumns").append($('<div><b>Metrics</b></div>'));
                        }
                        $("#ulmetriccolumns").append($('<li class="chartexpo_GoogleSheetAddon_tileMenu_li" originaltext="' + dataSourceDuplicateColumns[counter].ColumnName + '">' + dataSourceDuplicateColumns[counter].ColumnName + '</li>'));
                    }
                }
            }//end of for loop.
        }

        $(".chartexpo_GoogleSheetAddon_tileMenu_li").on('click', function () {
            var selectedValue = $(this).text();
            //alert(".chartexpo_GoogleSheetAddon_tileMenu_li selectedValue = > " + selectedValue);

            hideSearchMenu();
            if (isSearchMenuOpenedFromDimensionAddButton) {
                addDimensionsDropDownList(selectedValue); //If value is selected after add button click.
            }
            else if (isSearchMenuOpenedFromMeasureAddButton) {
                addMeasuresDropDownList(selectedValue);
            }
            else //If value is selected after dropdown menu click.
            {
                if (searchMenuSelectedObject != null) {
                    searchMenuSelectedObject.text(selectedValue);
                    searchMenuSelectedObject = null;
                }
                onSearchMenuColumnChange(selectedValue);
            }

            if (chartRequiredNoOfDimAndMetricsChosen()) {
                $("#divDrawChart").addClass('activeDrawButton');
                $('#divDrawChart').css("color", "#F37A2D");
                $('#divDrawChart').css("background-color", "white");
                $("#startRowTextBox").prop("readonly", false);
                $("#endRowTextBox").prop("readonly", false);
            }
            else {
                $("#divDrawChart").removeClass('activeDrawButton');
                $('#divDrawChart').css("color", "#B8B8B8");
                $("#startRowTextBox").prop("readonly", true);
                $("#endRowTextBox").prop("readonly", true);
            }
            var dimensionsInputLength = $('.dimension').length;
            var allowedDimensions = getSelectedChartAllowedDimensions();

            var meausresInputLength = $('.metric').length;
            var allowedMetrics = getSelectedChartAllowedMeasures();

            if (dimensionsInputLength >= allowedDimensions) {
                $("#addDimensionClick").hide();
            }

            if (meausresInputLength >= allowedMetrics) {
                $("#addMeasureClick").hide();
            }

            bindDivsDropEvent();

            //console.log("syncSelectedChartAtServer called from $('#.chartexpo_GoogleSheetAddon_tileMenu_li').on('click' ");

            if (!isSampleDataClicked) {
                syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
            }

            event.stopPropagation();
        });
    }
}
function isFloat(val) {
    var floatRegex = /^-?\d+(?:[.,]\d*?)?$/;
    if (!floatRegex.test(val))
        return false;

    val = parseFloat(val);
    if (isNaN(val))
        return false;
    return true;
}
function getDimensionColumnIndex(selectedData_From_ActualSheetData) {
    var dimensionColumnIndex = [];
    if (selectedData_From_ActualSheetData.length > 0) {
        var getMetricColumnsIndexes = getMetricColumnIndex(selectedData_From_ActualSheetData);
        if (getMetricColumnsIndexes.length > 0) {
            dimensionColumnIndex = selectedData_From_ActualSheetData[0].map(function (d, i) {
                return i;
            }).filter(function (d) {
                return getMetricColumnsIndexes.indexOf(d) == -1;
            });
        }
    }
    return dimensionColumnIndex;
}
function getMetricColumnIndex(selectedData_From_ActualSheetData) {
    var measureColumnIndex = [];
    if (selectedData_From_ActualSheetData.length > 0) {
        var tempArry = [];
        var counts = {};
        for (var column = 0; column < selectedData_From_ActualSheetData[0].length; column++) {
            tempArry = [];
            counts = {};
            for (var row = 0; row < selectedData_From_ActualSheetData.length; row++) {
                var data = selectedData_From_ActualSheetData[row][column];
                if (data != "") {
                    var t = isFloat(data.replace(/ /g, '').replace(/[$£€%Kr,]/g, ""));
                    tempArry.push(t);
                    //var patt = new RegExp("[0-9]");
                    //var res = patt.test(data);
                    //tempArry.push(res);
                }
            }
            for (var i = 0; i < tempArry.length; i++) {
                if (!counts.hasOwnProperty(tempArry[i])) {
                    counts[tempArry[i]] = 1;
                }
                else {
                    counts[tempArry[i]]++;
                }
            }
            if (tempArry.length > 0) {
                var count_true = counts[true] || 0;
                var count_false = counts[false] || 0;
                if (count_true > count_false) {// || tempArry.indexOf(false) == -1
                    measureColumnIndex.push(column);
                }
            }
        }
    }
    return measureColumnIndex;
}
function getDimensionColumns(selectedData_From_ActualSheetData) {
    var dimensionColumns = [];
    var dimensionColumnIndex = getDimensionColumnIndex(selectedData_From_ActualSheetData);
    if (dimensionColumnIndex.length > 0) {
        for (var i = 0; i < dimensionColumnIndex.length; i++) {
            if (dataSourceColumnsAll[0][dimensionColumnIndex[i]] != "") {
                dimensionColumns.push(dataSourceColumnsAll[0][dimensionColumnIndex[i]]);
            }
        }
    }
    return dimensionColumns;
}
function getMetricColumns(selectedData_From_ActualSheetData) {
    var metricColumns = [];
    var metricColumnIndex = getMetricColumnIndex(selectedData_From_ActualSheetData);
    if (metricColumnIndex.length > 0) {
        for (var i = 0; i < metricColumnIndex.length; i++) {
            if (dataSourceColumnsAll[0][metricColumnIndex[i]] != "") {
                metricColumns.push(dataSourceColumnsAll[0][metricColumnIndex[i]]);
            }
        }
    }
    return metricColumns;
}
function updateDataSourceColumnsAll() {
    for (var i = 0; i < dataSourceDuplicateColumns.length; i++) {
        var index = dataSourceColumnsAll[0].indexOf(dataSourceDuplicateColumns[i].ColumnName);
        if (index > -1) {
            dataSourceColumnsAll[0].splice(index, 1);
            dataSourceColumnsAll[0].splice(dataSourceDuplicateColumns[i].Col - 1, 0, dataSourceDuplicateColumns[i].ColumnName);
        }
    }
}

var previousDatasourceColumnValue;

function onDatasourceColumnBeforeChange() {
    previousDatasourceColumnValue = $(this).val();
}
var previousSelectedColumnValueIfFocusNotSet = '';

function onSearchMenuColumnChange(selectedDropdownValue) {
    //removing old element from selectedChartColumns.
    removeElementFromSelectedCollection(searchMenuPreviousSelectedValue);
    //previousDatasourceColumnValue = '';        
    //previousSelectedColumnValueIfFocusNotSet = selectedDropdownValue;
    //adding new column in selectedChartColumns.
    selectedChartColumns.push(selectedDropdownValue);

    //get A1Notation information.
    var columnA1NotationInformation = getColumnA1NotationFromSheetData(selectedDropdownValue);
    if (columnA1NotationInformation != "") {
        updateA1NotationColumn("", columnA1NotationInformation);
        //highlightSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumns);
        updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
    }
    //updating dropdown list.
    //updateDropdownList();
}

function getChartNameWithEllipses(selectedchartDisplayName) {
    if (selectedchartDisplayName.length > 23) {
        var sub = selectedchartDisplayName.substring(0, 23);
        return sub + "...";
    }
    else {
        return selectedchartDisplayName;
    }
}

function loadSheets() {
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    sheetNames = [];
    $('#dropdownSheets').empty();
    $('#dropdownSheets').append($('<option value="Select Sheet">Select Sheet</option>'));
    google.script.run.withSuccessHandler(function (sheetsName) {
        if (sheetsName != null && sheetsName != undefined) {
            sheetNames = sheetsName;
            // populate dropdownlist with sheetsName

            for (var counter = 0; counter < sheetsName.length; counter++) {
                $('#dropdownSheets').append($('<option value="' + sheetsName[counter] + '">' + sheetsName[counter] + '</option>'));
            } //end of for loop.
            var dropdownSheets = document.getElementById('dropdownSheets');
            dropdownSheets.addEventListener("change", onSheetSelectionChange);

            // On chartSelection by default current active sheetName selected in dropdownlist
            if (enableCurrentActiveSheetSelection) {
                if (syncMode == "Add") {
                    google.script.run.withSuccessHandler(function getCurrentActiveSheetDetail(SheetObj) {
                        if (SheetObj != null && SheetObj != "Select Sheet") {
                            onSheetSelectionChange(SheetObj);
                            $("#dropdownSheets").val(SheetObj);
                        }
                        else {
                            disableControl();
                        }
                    })
                        .withFailureHandler(
                        function (msg, element) {
                            handleError(msg);
                        }
                        ).getCurrentSheetName();
                }
            }
            $("#txtBoxHeaderRow").val(1);
            $('#chkHeaderRow').removeAttr("disabled");
            $('#chkHeaderRow').prop("checked", "checked");
            $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-active.png");
            if (!enableCurrentActiveSheetSelection) {
                $(".se-pre-con").fadeOut("slow");
            }
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).readAllSheetsNames();
}

function reloadSelectSheetDropdownListOnly() {
    var oldSelectedOption = $('#dropdownSheets').val();
    sheetNames = [];
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    google.script.run.withSuccessHandler(function (sheetsName) {
        if (sheetsName != null && sheetsName != undefined) {
            $('#dropdownSheets').empty();
            $('#dropdownSheets').append($('<option value="Select Sheet">Select Sheet</option>'));

            sheetNames = sheetsName;

            for (var counter = 0; counter < sheetsName.length; counter++) {
                if (oldSelectedOption == sheetsName[counter]) {
                    $('#dropdownSheets').append($('<option selected value="' + sheetsName[counter] + '">' + sheetsName[counter] + '</option>'));
                }
                else {
                    $('#dropdownSheets').append($('<option value="' + sheetsName[counter] + '">' + sheetsName[counter] + '</option>'));
                }
            } //end of for loop.
            var dropdownSheets = document.getElementById('dropdownSheets');
            dropdownSheets.addEventListener("change", onSheetSelectionChange);
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).readAllSheetsNames();
}

function onSheetSelectionChange(sheetName) {
    isMySavedChartExists = true;
    boldSelectedTopMenuOption("Bold", "normal", "normal", "normal", "normal", "normal");
    if ($("#addMeasureClick").css("display") == "none" || $("#addDimensionClick").css("display") == "none") {
        $("#addMeasureClick").css("display", "block");
        $("#addDimensionClick").css("display", "block");
    }
    $("#addMeasureClickContainer").css("margin-left", "0px");
    $("#addDimensionClickContainer").css("margin-left", "0px");
    var selectedSheetName = $(this).attr('id') == undefined ? sheetName : $(this).val();

    //sheet selected value is changed.
    if ($(this).attr('id') != undefined) {
        $('#chkHeaderRow').prop('checked', 'checked');
        $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-active.png");
        $("#txtBoxHeaderRow").val(1);
    }

    $('#dropDownDimensions0').empty();
    //$('#chkHeaderRow').prop("checked", "checked");
    onHeaderRowCheckboxChange($('#chkHeaderRow')[0]);
    selectedChartColumns = [];
    selectedChartA1NotationColumns = [];
    dataSourceColumns[0] = [];
    dataSourceColumnsAll[0] = [];
    dataSourceDuplicateColumns = [];
    actualSheetData = [];

    $('.metric').remove();
    //$('.dimension').not(':first').remove();
    $('.dimension').remove();
    clearTimeout(newchart_added_updated_timer_handler);

    if (selectedSheetName != "Select Sheet") {
        //selectedChartColumns = [];
        //dataSourceColumns[0] = [];
        //dataSourceColumnsAll[0] = [];
        //When open DataSource Screen and add new chartRecord syncMode is 'Add' but when change Sheet it changed to Edit instead of retain 'Add' Mode so I change it manually
        // syncMode = "Add"; // temp commented
        $(".se-pre-con").fadeIn("slow");
        //changeActiveSheet(selectedSheetName);
        //console.log("syncMode in onSheetSelectionChange(sheetName)" + syncMode);
        clearTimeout(syncSheetDataWithAddonTimerHandler);
        changeActiveSheetForSync(selectedSheetName);
        //loadSelectedSheetA1Notation(selectedSheetName);
        //logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-SelectSheetsDDL_Changed", "Charts");
        //enableControl();
        //$("#divDrawChart").addClass('activeDrawButton');
        //syncChart();
        //syncSheetDataWithAddonTimerHandler = setInterval(syncSheetDataWithGooglesheetAddonHandler, syncSheetDataWithAddonTimeSpan);
        $('.rearrangeText').css("z-index", "0");
    }
    else {
        $('.rearrangeText').css("z-index", "100");
        disableControl();
    }
}

function syncChart() {
    if (chartInEditModeSyncTime) return;
    var storedObjectInServerTempStorage = {
        selectedChart: selectedChart,
        selectedChartDisplayName: '',
        editableChartCustomName: '',
        headerRow: '',
        dataRows: '',
        myChart: '',
        editableChartGuid: '',
        dimension: '',
        defaultProperties: '',
        selectedChartCategory: '',
        openDataViewerOnLoad: '',
        sameContractList: '',
        chartAddedUpdatedIntoMyChartList: '',

        a1NotationInformation: '',
        useHeaderRow: $('#chkHeaderRow').prop("checked"),
        headerRowAnnotation: '',
        dataRowsAnnotation: '',
        fileName: '',
        sheetName: '',
        fileId: '',
        sheetId: '',
        headerRowNumber: headerRowNumber,
        dataRowFrom: '',
        dataRowTo: '',
        sheetWholedataRows: ''
    };
    // "Add" or "Edit"
    // On every sheet change, add new entry in database

    if (syncMode == "Add") {
        if (!isSampleDataClicked) {
            syncSelectedChartAtServer(selectedchartDisplayName, "Add", chartGuid, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());//storedObjectInServerTempStorage);
        }
    }
    else {
        storedObjectInServerTempStorage.editableChartGuid = chartGuid;
        //  console.log("syncSelectedChartAtServer called from syncChart() ");
        if (!isSampleDataClicked) {
            syncSelectedChartAtServer(selectedchartDisplayName, "Edit", chartGuid, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());//storedObjectInServerTempStorage);
        }
    }
}

function onHeaderChange() {
    //alert('header row is changing');
    selectedChartColumns = [];
    selectedChartA1NotationColumns = [];
    dataSourceColumns[0] = [];
    dataSourceColumnsAll[0] = [];
    dataSourceDuplicateColumns = [];
    sheetRecords = [];
    $('#dropDownDimensions0').empty();
    var headerRow = $("#txtBoxHeaderRow").val();
    headerRowNumber = parseInt(headerRow) - 1;
    //console.log("header row number = " + headerRowNumber);
    if (headerRow != null && headerRow != "" && actualSheetData != null) {
        headerRow = parseInt(headerRow) - 1;
        if (actualSheetData.length > headerRow) {
            //console.log("header row number inner = " + headerRowNumber);
            //dataSourceColumns[0] = removeEmptyColumns(actualSheetData[headerRow]);
            //dataSourceColumnsAll[0] = actualSheetData[headerRow];
            fillDatasourceColumnsAndSheetData(actualSheetData);
            if (dataSourceColumns[0].length == 0) {
                showMessageDialog("Empty Header", "Header row is empty. Please create logical column names", "confirmation", false, [], true);
            }
            //loading searchable menu.
            loadOriginalColumns();
            //loadDimensions('dropDownDimensions0');
            //attaching change event with dropdown list.
            attachChangeEvent('dropdownchangeclass');

        } // end of if statement.
        else {
            showMessageDialog("Empty Header", "Header row is empty. Please create logical column names", "confirmation", false, [], true);
        }
    } //end of outer if statement.
    updateSliderValue();
    $('.metric').remove();
    //$('.dimension').not(':first').remove();
    $('.dimension').remove();
    $(".dimension").unbind("click");
    $(".dimension").on('click', function () {
        showSearchMenuOnDropdownClick(this, event);
    });
    $(".dropdownchangeclass").on('focus', function () {
        $(this).attr("disabled", "true");
        showSearchMenuOnDropdownClick(this.parentElement);
    });
    if ($("#addMeasureClick").css("display") == "none" || $("#addDimensionClick").css("display") == "none") {
        $("#addMeasureClick").css("display", "block");
        $("#addDimensionClick").css("display", "block");
    }
    $("#addMeasureClickContainer").css("margin-left", "0px");
    $("#addDimensionClickContainer").css("margin-left", "0px");

    // console.log("syncSelectedChartAtServer called from onHeaderChange() {");
    if (!isSampleDataClicked) {
        syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
    }
}

function loadSelectedSheetA1Notation(sheetName) {
    // Get A1 notation of selected sheet.
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    google.script.run.withSuccessHandler(function (sheetA1Notation) {
        if (sheetA1Notation != null && sheetA1Notation != undefined && sheetA1Notation.length > 0) {
            sheetA1NotationDetails = sheetA1Notation;
            loadSheetData(sheetName);
            //console.log("R1C1 Notations=> " + JSON.stringify(sheetA1NotationDetails));
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }).getSelectedSheetCellsA1Notation(sheetName);
}

function getColumnA1NotationFromSheetData(dimensionName) {
    // SheetColumnIndex ,ColumnName
    for (var counter = 0; counter < datasourceColumnsWithIndex.length; counter++) {
        if (dimensionName == datasourceColumnsWithIndex[counter].ColumnName) {
            return datasourceColumnsWithIndex[counter].A1NotationInRCFormat;
        }//end of if statement.
    }//end of for loop.
    return "";
}

function updateA1NotationColumn(startValue, newA1Notation) {
    var isNewEntry = true;
    // for (var counter = 0; counter < selectedChartA1NotationColumns.length; counter++)
    // {
    //   if (selectedChartA1NotationColumns[counter].startsWith(startValue))
    //  {
    // selectedChartA1NotationColumns[counter] = newA1Notation;
    //        isNewEntry = false;
    //    } //end of if statement.
    //}//end of for loop.

    if (isNewEntry) {
        selectedChartA1NotationColumns.push(newA1Notation);
    }
}

function loadSheetData(sheetName) {
    //$(".se-pre-con").fadeIn("slow");
    //Get data of selected sheet.
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    google.script.run.withSuccessHandler(function (sheetData) {
        if (sheetData != null && sheetData != undefined && sheetData.length > 0) {
            //Create global variable and then work on this selected data.
            sheetData = JSON.parse(sheetData);
            actualSheetData = sheetData;
            fillDatasourceColumnsAndSheetData(sheetData);
            //loading searchable menu.
            loadOriginalColumns();
            //loading first dimension dropdown list.
            //loadDimensions('dropDownDimensions0');
            attachChangeEvent('dropdownchangeclass');
            //updateSliderValue();

            //If it is edit mode then load other dimensions and measures only on first time on loading.
            if (isEditModeClicked == true && isEditModeLoadingFromMyChartClick == true) {

                getHeaderNewColumns();

                var delayInMilliseconds = 3000; //3 second
                setTimeout(function () {
                    //loading other dimensions and measures.                                        
                    loadDimensionMeasuresInEditMode();
                    isEditModeLoadingFromMyChartClick = false;
                    $(".se-pre-con").fadeOut("slow");
                }, delayInMilliseconds);
            }
            else {
                updateSliderValue();
                $(".se-pre-con").fadeOut("slow");
            }

            //setting range mid value otherwise in UI it is shown at 0 on sheet loading.
            //$('#sheetDataRangeEnd').attr('value', Math.ceil(sheetRecords.length / 2));
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }).readDataFromSheetByName(sheetName);
}
//cvt = function (n) { return (String.fromCharCode(n + 'A'.charCodeAt(0) - 1)) }
// todo: null or empty check on settings column
function CheckSettingsColumnExistInSheetColumns(settingsColumn) {
    var isColumnMatched = false;
    for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
        if (settingsColumn.toString().toLowerCase() == dataSourceColumns[0][counter].toString().toLowerCase()) {
            isColumnMatched = true;
            break;
        }
    } //end of for loop.
    if (!isColumnMatched) {
        return false;
    }
    return true;
}

function matchHeaderColumnsInUseHeaderMode() {
    var savedHeaderColumns = getDateFromTempStorage("headerRow");
    savedHeaderColumns = JSON.parse(savedHeaderColumns);
    if (savedHeaderColumns.length > 0) {
        savedHeaderColumns = savedHeaderColumns[0];
        var isColumnMatched = false;
        for (var headerColumnsCounter = 0; headerColumnsCounter < savedHeaderColumns.length; headerColumnsCounter++) {
            isColumnMatched = false;
            for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
                if (savedHeaderColumns[headerColumnsCounter].toString().toLowerCase() == dataSourceColumns[0][counter].toString().toLowerCase()) {
                    isColumnMatched = true;
                    break;
                }
            } //end of for loop.
            if (!isColumnMatched) {
                return false;
            }
        } //end of for loop.
        return true;
    }
    return false;
}

// this method only fills global variable's values like dataSourceColumns,sheetA1Notation etc. but no effect on UI
function fillDatasourceColumnsAndSheetData(sheetData) {
    //debugger;
    //alert(JSON.stringify(dataSourceColumns[0]));
    //sheetData = [["Topic", "Sub topic-1", "Sub topic-2", "Count"],["Daily Food Sales", "Breakfast", "Waffles", "80"], ["Daily Food Sales", "Breakfast", "Eggs", "60"], ["Daily Food Sales", "Breakfast", "Pancakes", "45"], ["Daily Food Sales", "Breakfast", "Tea", "30"], ["Daily Food Sales", "Lunch", "Salid", "100"], ["Daily Food Sales", "Lunch", "Sandwich", "80"], ["Daily Food Sales", "Lunch", "Soup", "50"], ["Daily Food Sales", "Lunch", "Pie", "35"]];
    var headerRow = $("#txtBoxHeaderRow").val();
    dataSourceColumns[0] = [];
    dataSourceColumnsAll[0] = [];
    dataSourceDuplicateColumns = [];
    sheetRecords = [];
    headerRow = parseInt(headerRow) - 1;

    var headerColumns = [];
    var sheetA1Notation = '';
    //headerRow will be - 1 when $("#txtBoxHeaderRow").val() = 0
    for (var counter = headerRow; counter < sheetData.length; counter++) {

        if (counter == -1)//If there is no header column.
        {
            //console.log("sheet data" + JSON.stringify(sheetData[0]));
            var rowObject = getNonEmptyRowFromSheetData(sheetData);

            var sheetA1NotationWithNoHeader = [];
            if (rowObject != null && rowObject.rowData != null) {
                var nonEmptyRow = rowObject.rowData;
                for (var index = 0; index < nonEmptyRow.length; index++) {
                    if (nonEmptyRow[index] != null && nonEmptyRow[index] != "") {
                        headerColumns.push("Column " + convertToNumberingScheme((parseInt(index) + 1)));
                    }
                    else {
                        headerColumns.push("");
                    }
                    sheetA1NotationWithNoHeader.push("R" + rowObject.rowNumber + "C" + (parseInt(index) + 1));
                }//end of for loop.
                sheetA1Notation = sheetA1NotationWithNoHeader;
                dataSourceColumns[0] = removeEmptyColumns(headerColumns, sheetA1NotationWithNoHeader); //Adding header row.
                dataSourceColumnsAll[0] = headerColumns;
            }
        }
        else {
            var rowData = sheetData[counter];
            if (rowData != null) {
                rowData = rowData.join('');
                //if (rowData != "") {
                if (counter == parseInt(headerRow)) {
                    sheetA1Notation = '';
                    if (sheetA1NotationDetails.length > counter) {
                        sheetA1Notation = sheetA1NotationDetails[counter];
                    }
                    dataSourceColumns[0] = removeEmptyColumns(sheetData[counter], sheetA1Notation); //Adding header row.
                    dataSourceColumnsAll[0] = sheetData[counter];
                }
                else {
                    sheetRecords.push(sheetData[counter]);
                }
                //}
            }
        }
    }//end of for loop.
    setSelectRowRangeTextboxInitialValues(1, sheetRecords.length);
}

function removeSelectedDimension() {
    //alert('dimension remove clicked.');
    if (isSampleDataClicked)
        return;
    var objId = $(this).attr('id');

    var dropdownElementValue = $(this).parent().find('div').find('div').text();
    removeElementFromSelectedCollection(dropdownElementValue);
    hideTooltip();
    $(this).parent().remove();
    $(".dimension").each(function (index) {
        if (index % 2 == 0)
            $(this).css("margin-left", "0px");
        else
            $(this).css("margin-left", "7px");
    });

    $('#addDimensionClick').css('background', 'white');

    var dimensionsInputLength = $('.dimension').length;
    var isShowLeftMargin = dimensionsInputLength % 2;
    if (isShowLeftMargin == 0) {
        $('#addDimensionClickContainer').css('margin-left', '0px');
    }
    else {
        $('#addDimensionClickContainer').css('margin-left', '7px');
    }
    //updating dropdown list.
    //updateDropdownList();

    if (chartRequiredNoOfDimAndMetricsChosen()) {
        $("#divDrawChart").addClass('activeDrawButton');
        $('#divDrawChart').css("color", "#F37A2D");
        $('#divDrawChart').css("background-color", "white");
        $("#startRowTextBox").prop("readonly", false);
        $("#endRowTextBox").prop("readonly", false);
    }
    else {
        $("#divDrawChart").removeClass('activeDrawButton');
        $('#divDrawChart').css("color", "#B8B8B8");
        $("#startRowTextBox").prop("readonly", true);
        $("#endRowTextBox").prop("readonly", true);
    }

    //console.log("syncSelectedChartAtServer called from oremoveSelectedDimension() {");
    if (!isSampleDataClicked) {
        syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
    }
    $("#addDimensionClick").show();
}

function updateDropdownList() {
    //update other dropdownlists.
    //if (source == 'dimension') {
    $(".dimension select").each(function () {
        //var dropDownListId = $(this).attr('id');
        var dropDownListValue = $(this).val();

        $(this).empty();
        if (dataSourceColumns.length > 0) {
            for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
                if (!isColumnExistInSelectedCollection(dataSourceColumns[0][counter]))//If column does not exist in selectedChartColumns collection then add it.
                {
                    $(this).append($('<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option>'));
                }
            }
        }//end of if statement.
        if (dropDownListValue != null && dropDownListValue != '') {
            $(this).append($('<option value="' + dropDownListValue + '">' + dropDownListValue + '</option>'));
            var optionElements = $(this).find('option').length;
            $(this).prop('selectedIndex', (optionElements - 1));
        }
    });
    //}
    //else {
    $(".metric select").each(function () {
        var dropDownListValue = $(this).val();
        $(this).empty();
        if (dataSourceColumns.length > 0) {
            for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
                if (!isColumnExistInSelectedCollection(dataSourceColumns[0][counter]))//If column does not exist in selectedChartColumns collection then add it.
                {
                    $(this).append($('<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option>'));
                }
            }
        }//end of if statement.
        if (dropDownListValue != null && dropDownListValue != '') {
            $(this).append($('<option value="' + dropDownListValue + '">' + dropDownListValue + '</option>'));
            var optionElements = $(this).find('option').length;
            $(this).prop('selectedIndex', (optionElements - 1));
        }
    });
    //}
}

function removeSelectedMeasure() {
    //alert('measure remove clicked.');
    if (isSampleDataClicked)
        return;
    var objId = $(this).attr('id');
    hideTooltip();
    var dropdownElementValue = $(this).parent().find('div').find('div').text();
    removeElementFromSelectedCollection(dropdownElementValue);

    $(this).parent().remove();
    $(".metric").each(function (index) {
        if (index % 2 == 0)
            $(this).css("margin-left", "0px");
        else
            $(this).css("margin-left", "7px");
    });

    $('#addMeasureClick').css('background', 'white');

    var metricsInputLength = $('.metric').length;
    var isShowLeftMargin = metricsInputLength % 2;
    if (isShowLeftMargin == 0) {
        $('#addMeasureClickContainer').css('margin-left', '0px');
    }
    else {
        $('#addMeasureClickContainer').css('margin-left', '7px');
    }
    //updating dropdown list.
    //updateDropdownList();

    if (chartRequiredNoOfDimAndMetricsChosen()) {
        $("#divDrawChart").addClass('activeDrawButton');
        $('#divDrawChart').css("color", "#F37A2D");
        $('#divDrawChart').css("background-color", "white");
        $("#startRowTextBox").prop("readonly", false);
        $("#endRowTextBox").prop("readonly", false);
    }
    else {
        $("#divDrawChart").removeClass('activeDrawButton');
        $('#divDrawChart').css("color", "#B8B8B8"); // To do, manage it from class.
        $("#startRowTextBox").prop("readonly", true);
        $("#endRowTextBox").prop("readonly", true);
    }
    //console.log("syncSelectedChartAtServer called from removeSelectedMeasure() {");
    if (!isSampleDataClicked) {
        syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
    }
    $("#addMeasureClick").show();
}

function onDatasourceColumnChange() {
    //removing old element from selectedChartColumns.
    removeElementFromSelectedCollection((previousDatasourceColumnValue == '' ? previousSelectedColumnValueIfFocusNotSet : previousDatasourceColumnValue));
    previousDatasourceColumnValue = '';
    var selectedDropdownValue = $(this).val();
    previousSelectedColumnValueIfFocusNotSet = selectedDropdownValue;
    //adding new column in selectedChartColumns.
    selectedChartColumns.push(selectedDropdownValue);

    //get A1Notation information.
    var columnA1NotationInformation = getColumnA1NotationFromSheetData(selectedDropdownValue);
    if (columnA1NotationInformation != "") {

        updateA1NotationColumn("", columnA1NotationInformation);

        // console.log("onDatasourceColumnChange");
        updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
    }
    //updating dropdown list.
    //updateDropdownList();
}

function isColumnExistInSelectedCollection(columnName) {
    for (var counter = 0; counter < selectedChartColumns.length; counter++) {
        if (columnName != null && columnName == selectedChartColumns[counter]) {
            return true;
        }//end of if statement.
    }//end of for loop.
    return false;
}

function removeElementFromSelectedCollection(columnName) {
    for (var counter = 0; counter < selectedChartColumns.length; counter++) {
        if (columnName != null && columnName == selectedChartColumns[counter]) {
            selectedChartColumns.splice(counter, 1);
            if (selectedChartA1NotationColumns.length > counter) {
                selectedChartA1NotationColumns.splice(counter, 1);
                //var columnA1NotationInformation = getColumnA1NotationFromSheetData(columnName);
                //removeFromA1NotationColumns(columnA1NotationInformation);
                //console.log("removeElementFromSelectedCollection");
                updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
            }
            break;
        }//end of if statement.
    }//end of for loop.
}

function removeFromA1NotationColumns(a1Notation) {
    for (var counter = 0; counter < selectedChartA1NotationColumns.length; counter++) {
        if (selectedChartA1NotationColumns[counter].startsWith(a1Notation)) {
            selectedChartA1NotationColumns.splice(counter, 1);
            break;
        }//end of if statement.

    }//end of for loop.
}

function attachClickEvent(className, bindingSource) {
    var imgList;
    // get all the elements with className 'btn'. It returns an array
    var imgList = document.getElementsByClassName(className);
    // get the lenght of array defined above
    var listLength = imgList.length;
    var i = 0;
    // run the for look for each element in the array
    for (; i < listLength; i++) {
        // attach the event listener
        if (bindingSource == 'dimension')
            imgList[i].addEventListener("click", removeSelectedDimension);
        else
            imgList[i].addEventListener("click", removeSelectedMeasure);
    }
}
function getSameContractCharts(minDimension, minMetric, ruleObject) {
    var chartDimensionArray = [];
    for (var i = 0; i < chartRulesObject.length; i++) {
        for (var j = 0; j < ruleObject.similarCharts.length; j++) {
            if (ruleObject.similarCharts[j] == chartRulesObject[i].ChartName) {
                chartDimensionArray.push(chartRulesObject[i]);
            }
        }
    }
    return chartDimensionArray;
}
function attachChangeEvent(className) {
    var dropdownList;
    // get all the elements with className 'btn'. It returns an array
    var dropdownList = document.getElementsByClassName(className);
    // get the lenght of array defined above
    var listLength = dropdownList.length;
    var i = 0;
    // run the for look for each element in the array
    for (; i < listLength; i++) {
        // attach the event listener
        //if (bindingSource == 'dimension')
        dropdownList[i].addEventListener("change", onDatasourceColumnChange);
        dropdownList[i].addEventListener("focus", onDatasourceColumnBeforeChange);
    }
}

function addDimensionsDropDownListForSampleData() {
    if (dataSourceColumns.length > 0) {

        var dimensionsInputLength = $('.dimension').length;
        var allowedDimensions = getSelectedChartAllowedDimensions();

        if (allowedDimensions == 0) {
            showMessageDialog("Rule Missing", "Please provide maximum allowed dimension(s) for this chart in rules!", "confirmation", false, [], true);
            return;
        }
        if (dimensionsInputLength >= allowedDimensions) {
            //showMessageDialog("Maximum Limit Reached", "Dimension(s) maximum limit for this chart is reached!", "confirmation", false, [],true);
            return;
        }

        var leftSpace = ' style= margin-bottom:4px;';
        var isShowLeftMargin = dimensionsInputLength % 2;
        if (isShowLeftMargin != 0)
            leftSpace += 'margin-left:7px;';
        var dimensionsDropdownList = '<div class="dimension" draggable="true" ' + leftSpace + '><select id="dropDownDimensions' + dimensionsInputLength + '" class="DropdownList dropdowncolor dropdownchangeclass" style="width:110px;-webkit-appearance:none;padding-left:4px;">';

        var firstAddedColumn = '';
        for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
            if (!isColumnExistInSelectedCollection(dataSourceColumns[0][counter]))//If column does not exist in selectedChartColumns collection then add it.
            {
                if (firstAddedColumn == '')
                    firstAddedColumn = dataSourceColumns[0][counter];
                dimensionsDropdownList = dimensionsDropdownList + '<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option >';
            } //end of if statement.
        } //end of for loop.

        if (firstAddedColumn == '') {
            //alert('No more dimension/measure exist!');
            showMessageDialog("Dimension/Measure", "No more dimension/measure exist!", "confirmation", false, [], true);
            return;
        }
        //Updating selectedChartColumns list.
        if (firstAddedColumn != '') {
            selectedChartColumns.push(firstAddedColumn);
            //get A1Notation information.
            var columnA1NotationInformation = getColumnA1NotationFromSheetData(firstAddedColumn);
            if (columnA1NotationInformation != "") {
                //selectedChartA1NotationColumns.push(columnA1NotationInformation);
                updateA1NotationColumn("", columnA1NotationInformation);
                //highlightSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumns);
                //console.log("addDimensionsDropDownList");
                updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
            }
        }

        dimensionsDropdownList = dimensionsDropdownList + '</select><img id="dimensionRemove' + dimensionsInputLength + '" style="margin-left:5px;" class="imagecontainer dimensionRemoveClass" src="https://apps.polyvista.com/GooglesheetFeb2021/Scripts/Polyvista/feedback177/icons/remove-selection.svg" title="remove" /></div>';

        //$('#divDimensionsContainer').append($(dimensionsDropdownList));
        $('#addDimensionClickContainer').before($(dimensionsDropdownList));
        attachClickEvent('dimensionRemoveClass', 'dimension');
        attachChangeEvent('dropdownchangeclass');
        bindDivsDropEvent();
    }
}

function getSelectedChartAllowedDimensions() {
    for (var counter = 0; counter < chartRulesObject.length; counter++) {
        if (selectedChartNameFromSelectChartUI != "" && chartRulesObject[counter].ChartName == selectedChartNameFromSelectChartUI)
            return chartRulesObject[counter].MaxDim;
    }
    return 0;
}

function getSelectedChartAllowedMeasures() {
    for (var counter = 0; counter < chartRulesObject.length; counter++) {
        if (selectedChartNameFromSelectChartUI != "" && chartRulesObject[counter].ChartName == selectedChartNameFromSelectChartUI)
            return chartRulesObject[counter].MaxMetric;
    }
    return 0;
}

function loadDimensions(dropDownListName) {

    $('#' + dropDownListName).empty();

    if (dataSourceColumns.length > 0) {
        for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
            $('#' + dropDownListName).append($('<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option>'));
        }
        //Updating selectedChartColumns list.
        if (dataSourceColumns[0].length > 0) {
            selectedChartColumns.push(dataSourceColumns[0][0]);
            //get A1Notation information.
            var columnA1NotationInformation = getColumnA1NotationFromSheetData(dataSourceColumns[0][0]);
            if (columnA1NotationInformation != "") {
                //selectedChartA1NotationColumns.push(columnA1NotationInformation);
                updateA1NotationColumn("", columnA1NotationInformation);
                //highlightSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumns);
                //console.log("loadDimensions");
                updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
            }
        }
    }//end of if statement.
    $(".dimension").unbind("click");
    $(".dimension").on('click', function () {
        showSearchMenuOnDropdownClick(this, event);
    });
    $(".dropdownchangeclass").on('focus', function () {
        $(this).attr("disabled", "true");
        showSearchMenuOnDropdownClick(this.parentElement);
    });
}

function addMeasuresDropDownList(columnSelected, isSettingsColumnExistInSheet) {
    if (dataSourceColumns.length > 0) {
        var deletedColumnColor = "";
        var makeColumnDraggable = 'draggable="true"';
        if (isSettingsColumnExistInSheet != null && isSettingsColumnExistInSheet == false) {
            deletedColumnColor = "background-color:#f08080;";
            makeColumnDraggable = "";
        }

        var threedigitsrandom = Math.floor(100 + Math.random() * 900);

        var meausresInputLength = $('.metric').length;
        var allowedMetrics = getSelectedChartAllowedMeasures();

        var webkitAprearance = '';
        if (isSampleDataClicked) {
            webkitAprearance = '-webkit-appearance:none';
        }

        var leftSpace = ' style=margin-bottom:4px;'
        var isShowLeftMargin = meausresInputLength % 2;
        if (isShowLeftMargin != 0) {
            leftSpace += 'margin-left:7px;'
            $('#addMeasureClickContainer').css('margin-left', '0px');
        }
        else {
            $('#addMeasureClickContainer').css('margin-left', '7px');
        }

        var measuresDropdownList = '<div column="metric" class="metric" ' + makeColumnDraggable + leftSpace + '><div id="dropDownMeasures' + meausresInputLength + threedigitsrandom + '" class="DropdownList metricdropdowncolor dropdownchangeclass" style="width:100px;padding-left:4px;;float:left;' + deletedColumnColor + '" isDeletedColumn=' + (isSettingsColumnExistInSheet == false ? "true" : "false") + '>';

        var firstAddedColumn = columnSelected;
        measuresDropdownList = measuresDropdownList + '<div style="padding-top:4px;float:left;width: 83px;overflow: hidden;white-space: nowrap;text-overflow:ellipsis;" value="' + columnSelected + '">' + columnSelected + '</div >';
        measuresDropdownList = measuresDropdownList + '<div style="float:right;padding-top:4px;padding-right:5px;"><img src="https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png"></div>';

        ////for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
        ////    if (!isColumnExistInSelectedCollection(dataSourceColumns[0][counter]))//If column does not exist in selectedChartColumns collection then add it.
        ////    {
        ////        if (firstAddedColumn == '')
        ////            firstAddedColumn = dataSourceColumns[0][counter];
        ////        measuresDropdownList = measuresDropdownList + '<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option >';
        ////    } //end of if statement.
        ////} //end of for loop.

        if (firstAddedColumn == '') {
            //alert('No more dimension/measure exist!');
            showMessageDialog("Dimension/Measure", "No more dimension/measure exist!", "confirmation", false, [], true);
            return;
        }

        //Updating selectedChartColumns list.
        if (firstAddedColumn != '') {
            selectedChartColumns.push(firstAddedColumn);
            //get A1Notation information.
            var columnA1NotationInformation = getColumnA1NotationFromSheetData(firstAddedColumn);
            if (columnA1NotationInformation != "") {
                updateA1NotationColumn("", columnA1NotationInformation);
                //console.log("addMeasuresDropDownList");
                updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
            }
        }
        measuresDropdownList = measuresDropdownList + '</div><img id="measureRemove' + meausresInputLength + '" style="margin-left:5px;padding-top:6px;" class="imagecontainer metricRemoveClass" src="https://apps.polyvista.com/GooglesheetFeb2021/Scripts/Polyvista/feedback177/icons/remove-selection.svg" title="remove"/></div>';

        $('#addMeasureClickContainer').before($(measuresDropdownList));
        attachClickEvent('metricRemoveClass', 'measure');
        attachChangeEvent('dropdownchangeclass');

        meausresInputLength = $('.metric').length;
        if (meausresInputLength == allowedMetrics) {
            $('#addMeasureClick').css('background', 'white');
        }

        $(".metric").unbind("click");
        $(".metric").on('click', function () {
            showSearchMenuOnDropdownClick(this, event);
        });
        $(".dropdownchangeclass").on('focus', function () {
            $(this).attr("disabled", "true");
            showSearchMenuOnDropdownClick(this.parentElement);
        });

        //set row slider to maximum value for the first time when dimension or measure is added.
        if (meausresInputLength == 1 && $('.dimension').length == 0) {
            updateSliderValue(0, sheetRecords.length);
        }
    }
}

function addMeasuresDropDownListForSampleData() {
    if (dataSourceColumns.length > 0) {
        var meausresInputLength = $('.metric').length;
        var allowedMetrics = getSelectedChartAllowedMeasures();

        if (allowedMetrics == 0) {
            showMessageDialog("Max Allowed Metric(s) Missing in Rules", "Please provide maximum allowed metric(s) for this chart in rules!", "confirmation", false, [], true);
            return;
        }
        if (meausresInputLength >= allowedMetrics) {
            //showMessageDialog("Metric(s) Maximum Limit Reached", "Metric(s) maximum limit for this chart is reached!", "confirmation", false, [],true);
            return;
        }

        var leftSpace = '';
        var isShowLeftMargin = meausresInputLength % 2;
        if (isShowLeftMargin != 0)
            leftSpace = ' style=margin-left:7px;margin-bottom:4px;'
        var measuresDropdownList = '<div class="metric" ' + leftSpace + '><select id="dropDownMeasures' + meausresInputLength + '" class="DropdownList metricdropdowncolor dropdownchangeclass" style="width:110px;-webkit-appearance:none;padding-left:4px;" disabled>';

        var firstAddedColumn = '';
        for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
            if (!isColumnExistInSelectedCollection(dataSourceColumns[0][counter]))//If column does not exist in selectedChartColumns collection then add it.
            {
                if (firstAddedColumn == '')
                    firstAddedColumn = dataSourceColumns[0][counter];
                measuresDropdownList = measuresDropdownList + '<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option >';
            } //end of if statement.
        } //end of for loop.

        if (firstAddedColumn == '') {
            //alert('No more dimension/measure exist!');
            showMessageDialog("Dimension/Measure", "No more dimension/measure exist!", "confirmation", false, [], true);
            return;
        }

        //Updating selectedChartColumns list.
        if (firstAddedColumn != '') {
            selectedChartColumns.push(firstAddedColumn);
            //get A1Notation information.
            var columnA1NotationInformation = getColumnA1NotationFromSheetData(firstAddedColumn);
            if (columnA1NotationInformation != "") {
                updateA1NotationColumn("", columnA1NotationInformation);
                //console.log("addMeasuresDropDownList");
                updateSelectedChartA1NotificationColumns(startSliderRange, endSliderRange);
            }
        }
        measuresDropdownList = measuresDropdownList + '</select><img id="measureRemove' + meausresInputLength + '" style="margin-left:5px" class="imagecontainer metricRemoveClass" src="https://apps.polyvista.com/GooglesheetFeb2021/Scripts/Polyvista/feedback177/icons/remove-selection.svg" title="remove"/></div>';

        $('#addMeasureClickContainer').before($(measuresDropdownList));
        attachClickEvent('metricRemoveClass', 'measure');
        attachChangeEvent('dropdownchangeclass');
        //updateDropdownList();
    }
}

function loadMeasures(dropDownListName) {
    if (dataSourceColumns.length > 0) {
        //$('#' + dropDownListName).append($('<option value="' + 'Please select' + '">' + 'Please select' + '</option>'));
        for (var counter = 0; counter < dataSourceColumns[0].length; counter++) {
            $('#' + dropDownListName).append($('<option value="' + dataSourceColumns[0][counter] + '">' + dataSourceColumns[0][counter] + '</option>'));
        }
    }//end of if statement.
}

function showTooltipOnCreateChartButton(chartRuleObject, currentElement) {
    var html = '<ul style="margin-top:-13px;">';
    html += '<p style="margin-left:-14px;margin-bottom:0px">Please provide the required fields:</p>';
    var dimension = chartRuleObject.DimensionText;
    var metric = chartRuleObject.MeasureText;
    var only = "Only ";
    var columnMetric = "column";
    var columnDimension = "column";
    if (chartRuleObject.MaxDim > 1) {
        dimension = chartRuleObject.DimensionText;
        only = "";
        columnDimension = "columns";
    }
    only = "Only ";
    column = "column";
    if (chartRuleObject.MaxMetric > 1) {
        metric = chartRuleObject.MeasureText;
        only = "";
        columnMetric = "columns";
    }

    if (chartRuleObject.MinMetric == chartRuleObject.MaxMetric) {
        html += '<li>'  + chartRuleObject.MinMetric + ' ' + columnMetric + ' for ' + metric + '</li>';
    }
    else if (chartRuleObject.MaxMetric > 1 && chartRuleObject.MaxMetric == 2)
    {
        html += '<li>' + chartRuleObject.MinMetric + ' or ' + chartRuleObject.MaxMetric + ' ' + columnMetric + ' for ' + metric + '</li>';
    }
    else {
        html += '<li>' + columnMetric + ' (>= ' + chartRuleObject.MinMetric + ' & <= ' + chartRuleObject.MaxMetric + ') for ' + metric + '</li>';
    }
    if (chartRuleObject.MinDim == chartRuleObject.MaxDim) {
        html += '<li>' + chartRuleObject.MinDim + ' ' + columnDimension + ' for ' + dimension + '</li>';
    }
    else if (chartRuleObject.MaxDim > 1 && chartRuleObject.MaxDim == 2)
    {
        html += '<li>' + chartRuleObject.MinDim + ' or ' + chartRuleObject.MaxDim + ' ' + columnDimension + ' for ' + dimension + '</li>';
    }
    else {
        html += '<li>' + columnDimension + ' (>= ' + chartRuleObject.MinDim + ' & <= ' + chartRuleObject.MaxDim + ') for ' + dimension + '</li>';
    }
    html += '<li>Row range</li>';
    html += '</ul>';
    generateTooltip(false, "createChart", html, event, currentElement);
}

function generateTooltip(isCloseButton, toolTipType, toolTipContent, event, currentElement) {
    isCloseButton = isCloseButton || false;
    toolTipContent = toolTipContent || "";
    toolTipType = toolTipType || "createChart";
    hideTooltip();
    var html = "";
    if (toolTipType.toLowerCase() == "createChart".toLowerCase()) {
        var x = $(currentElement).offset().left;
        var y = $(currentElement).offset().top;
        var columnDivHeight = $(currentElement).height();
        html += '<div class="toolTip-createChart">';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
        var tooltipContainerHeight = $('.toolTip-createChart').height();
        var topSpace = 11;//for show on exact location
        var leftSpace = 20;//for show on exact location
        $('.toolTip-createChart').css('left', x - leftSpace);
        $('.toolTip-createChart').css('top', y - tooltipContainerHeight - columnDivHeight - topSpace);
    }
    else if (toolTipType.toLowerCase() == "dimensions".toLowerCase()) {
        html += '<div class="toolTip-dimensions">';
        html += '<div class="triangle-up"></div>';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
        var position = $("#info-icon-dimension").offset();
        $(".toolTip-dimensions").css({ "top": position.top + 25 });
        $(".triangle-up").css({ "left": position.left - 15 });
    }
    if (toolTipType.toLowerCase() == "metrics".toLowerCase()) {
        html += '<div class="toolTip-metrics">';
        html += '<div class="triangle-up"></div>';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
        var position = $("#info-icon-metric").offset();
        $(".toolTip-metrics").css({ "top": position.top + 25 });
        $(".triangle-up").css({ "left": position.left - 15 });
    }
    else if (toolTipType.toLowerCase() == "help".toLowerCase()) {
        var x = $(currentElement).offset().left;
        var y = $(currentElement).offset().top;
        var columnDivHeight = $(currentElement).height();
        html += '<div class="toolTip-help">';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
        var tooltipContainerHeight = $('.toolTip-help').height();
        var leftSpace = 33;//for show on exact location
        $('.toolTip-help').css('left', x - 33);
        $('.toolTip-help').css('top', y - tooltipContainerHeight - columnDivHeight);
    }
    else if (toolTipType.toLowerCase() == "addSheetSampleData".toLowerCase()) {
        html += '<div class="toolTip-sheetSampleData">';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
    }
    else if (toolTipType.toLowerCase() == "headerRow".toLowerCase()) {
        html += '<div class="toolTip-headerRow">';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
    }
    else if (toolTipType.toLowerCase() == "deletedDimension".toLowerCase()) {
        var x = $(currentElement).offset().left;
        var y = $(currentElement).offset().top;
        var columnDivHeight = $(currentElement).height();
        html += '<div class="toolTip-deletedColumn">';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
        var tooltipContainerHeight = $('.toolTip-deletedColumn').height();
        $('.toolTip-deletedColumn').css('left', x);
        $('.toolTip-deletedColumn').css('top', y - tooltipContainerHeight - columnDivHeight);
    }
    else if (toolTipType.toLowerCase() == "deletedMetric".toLowerCase()) {
        var x = $(currentElement).offset().left;
        var y = $(currentElement).offset().top;
        var columnDivHeight = $(currentElement).height();
        html += '<div class="toolTip-deletedColumn">';
        html += toolTipContent;
        if (isCloseButton) {
            html += '<div class="tip-close">Close</div>';
        }
        html += '</div>';
        $(".info").append(html);
        var tooltipContainerHeight = $('.toolTip-deletedColumn').height();
        $('.toolTip-deletedColumn').css('left', x);
        $('.toolTip-deletedColumn').css('top', y - tooltipContainerHeight - columnDivHeight);
    }
    //$(".info").append(html);
    if (isCloseButton) {
        $(".tip-close").click(function () {
            hideTooltip();
        });
    }
}
var isOriginalTable = false;

// called in case of both Sample and Sheet data
// this method set dim/measures selected and get latest data from sheet and open chart viewer
function createChartWithLatestData() {
    //google.script.run.showChartViewerInDialog();
    var chartRuleObject = getSelectedChartRulesObject();
    selectedchartDisplayName = chartRuleObject.ChartDisplayName;
    if (chartRuleObject == null) {
        //alert("Please add rule for selected chart before drawing!");
        showMessageDialog("Rule Missing", "Please add rule for selected chart before drawing!", "confirmation", false, [], true);
        return;
    }
    var drawChartObject = { ChartName: '', Sheet: '', HeaderRow: 0, Dimensions: [], Measures: [], RowStartIndex: 0, RowLastIndex: 0 };
    var dimensions = [];
    var measures = [];
    drawChartObject.ChartName = selectedChartNameFromSelectChartUI;
    drawChartObject.Sheet = $('#dropdownSheets').val();
    drawChartObject.HeaderRow = parseInt($('#txtBoxHeaderRow').val());
    $(".dimension > div:first-child").each(function (index) {
        dimensions.push($(this).text());
    });
    drawChartObject.Dimensions = dimensions;
    $(".metric > div:first-child").each(function (index) {
        measures.push($(this).text());
    });
    drawChartObject.Measures = measures;

    if (isSampleDataClicked) {
        dimensions = [];
        measures = [];
        $(".dimension > select").each(function (index) {
            dimensions.push($(this).val());
        });
        $(".metric > select").each(function (index) {
            measures.push($(this).val());
        });

        drawChartObject.Dimensions = dimensions;
        drawChartObject.Measures = measures;
    }

    //drawChartObject.RowStartIndex = parseInt(document.getElementById("sheetDataRange").value);
    //drawChartObject.RowLastIndex = parseInt(document.getElementById("sheetDataRangeEnd").value);
    drawChartObject.RowLastIndex = maxValue;
    samecontractlist = getSameContractCharts(chartRuleObject.MinDim, chartRuleObject.MinMetric, chartRuleObject);
    //window.localStorage.setItem("sameContractList", JSON.stringify(samecontractlist));
    saveDateInTempStorage("sameContractList", JSON.stringify(samecontractlist))

    tableHeaderRow = [[]];
    var tableBodyRows = [];
    tableHeaderRow[0] = dimensions.concat(measures);
    if (isSampleDataClicked) { //getting sample data.
        //var dataRows = localStorage.getItem("dataRows");
        //console.log(JSON.stringify(tableHeaderRow[0]));
        var dataRows = getDateFromTempStorage("dataRows");
        var tableOriginalHeaderRow = [columnNameMapper[selectedChart]];
        if (dataRows != null && dataRows.length > 0) {
            tableBodyRows = JSON.parse(dataRows);
            //alert(JSON.stringify(tableBodyRows));
            logUserActionIntoDatabase(selectedChart + "-Chart_Drawn", "Charts");
            //alert(exploreSampleDataLinkClicked);
            if (exploreSampleDataLinkClicked) {
                //window.localStorage.setItem('openDataViewerOnLoad', "true");
                saveDateInTempStorage('openDataViewerOnLoad', "true");
                exploreSampleDataLinkClicked = false;
            }
            else {
                //window.localStorage.setItem('openDataViewerOnLoad', "false");
                saveDateInTempStorage('openDataViewerOnLoad', "false");
            }
            if (isOriginalTable == false) {
                tableOriginalBodyRows = tableBodyRows;
                isOriginalTable = true;
            }
            tblDraggableBodyContent = convertChartDataColumnsIntoDraggableFormat(tableHeaderRow, tableOriginalHeaderRow, tableOriginalBodyRows);
            var processedObject = getProcessedDataObjectForChartViewer("sampledata");
            processedObject.dataRows = tblDraggableBodyContent;
            processedObject.headerRow = tableHeaderRow;
            storeDataInLocalStorage(processedObject);
            openChartViewer("sampledata", tableHeaderRow, tblDraggableBodyContent, selectedChartNameFromSelectChartUI, selectedchartDisplayName);
        }
    }
    else {
        //var startRowIndex = parseInt(drawChartObject.RowStartIndex) == 0 ? 0 : parseInt(drawChartObject.RowStartIndex) - 1;

        //console.log("Min Value:" + minValue);
        //console.log("Max Value:" + drawChartObject.RowLastIndex);

        var headerRowNumber = parseInt($("#txtBoxHeaderRow").val());
        var startIndexNumber;
        if (headerRowNumber == 0) {
            startIndexNumber = startSliderRange;
        }
        else {
            startIndexNumber = startSliderRange;
        }
        //for (var counter = (headerRowNumber + minValue - 1); counter < (drawChartObject.RowLastIndex + headerRowNumber); counter++) {
        for (var counter = (startIndexNumber); counter <= (drawChartObject.RowLastIndex); counter++) {
            if (counter < 0) {
                counter = 0;
            }
            if (counter <= drawChartObject.RowLastIndex) { //if counter value is less than slider value.
                var filteredDataObject = getDataObjectWithFilteredColumns(dimensions, measures, counter);

                tableBodyRows.push(filteredDataObject);
            }
        }

        //console.log("data:" + JSON.stringify(tableBodyRows));
        //console.log("min value:" + minValue);
        //console.log("max value:" + maxValue);
        tableBodyRows = processEmptyRowCells(tableBodyRows, dimensions, measures);

        logUserActionIntoDatabase(selectedChart + "-DrawChartFromSheetData", "Charts");

        // if (isEditModeClicked) {

        // In case of autoSync, always open chart viewer in edit mode

        saveDateInTempStorage('headerRow', JSON.stringify(tableHeaderRow));
        saveDateInTempStorage('dataRows', JSON.stringify(tableBodyRows));

        saveDateInTempStorage('selectedDimensions', JSON.stringify(drawChartObject.Dimensions));
        saveDateInTempStorage('selectedMeasures', JSON.stringify(drawChartObject.Measures));

        //setting local storage.                
        saveDateInTempStorage('a1NotationInformation', oldSelectedCellsA1Notations);

        saveDateInTempStorage('useHeaderRow', $('#chkHeaderRow').prop("checked"));
        saveDateInTempStorage("fileName", fileName);
        saveDateInTempStorage("sheetName", sheetName);
        saveDateInTempStorage("fileId", fileId);
        saveDateInTempStorage("sheetId", sheetId);
        saveDateInTempStorage("headerRowNumber", $("#txtBoxHeaderRow").val());
        saveDateInTempStorage("dataRowFrom", minValue);
        saveDateInTempStorage("dataRowTo", maxValue);
        saveDateInTempStorage("headerRowAnnotation", JSON.stringify(headerRowAnnotation));
        saveDateInTempStorage("dataRowsAnnotation", JSON.stringify(dataRowsAnnotation));

        openChartViewer("editdata", tableHeaderRow, tableBodyRows, selectedChartNameFromSelectChartUI, selectedchartDisplayName);

        /*
    }
    else {
        openChartViewer("sheetdata", tableHeaderRow, tableBodyRows, selectedChartNameFromSelectChartUI, selectedchartDisplayName);
    }*/

        //$("#divDrawChart").css("cursor", "pointer");
    }
}

function getLatestA1NotationFromSheet(sheetName, obj) {
    //$(obj).css('cursor', 'wait');
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    $(".se-pre-con").fadeIn("slow");
    // Get A1 notation of selected sheet.
    google.script.run.withSuccessHandler(function (sheetA1Notation) {
        if (sheetA1Notation != null && sheetA1Notation != undefined && sheetA1Notation.length > 0) {
            sheetA1NotationDetails = sheetA1Notation;
            getLatestDataFromSheet(sheetName);
        }
    })
        .withFailureHandler(
        function (msg, element) {
            handleError(msg);
            //$(obj).css('cursor', 'pointer');
            $(".se-pre-con").fadeOut("slow");
        }).getSelectedSheetCellsA1Notation(sheetName);
}

function getLatestDataFromSheet(sheetName) {
    //Get data of selected sheet.
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    google.script.run.withSuccessHandler(function (sheetData) {
        if (sheetData != null && sheetData != undefined && sheetData.length > 0) {
            //Create global variable and then work on this selected data.
            sheetData = JSON.parse(sheetData);
            actualSheetData = sheetData;
            updateSheetRecordsWithLatestData(sheetData);
            //$(".se-pre-con").fadeOut("slow");
            //updateSliderValue(0, 1);
            createChartWithLatestData();
        }
        else {
            //$('#divDrawChart').css('cursor', 'pointer');
            $(".se-pre-con").fadeOut("slow");
        }
    })
        .withFailureHandler(
        function (msg, element) {
            //$('#divDrawChart').css('cursor', 'pointer');
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }).readDataFromSheetByName(sheetName);
}

function updateSheetRecordsWithLatestData(sheetData) {
    var headerRow = $("#txtBoxHeaderRow").val();
    sheetRecords = [];
    headerRow = parseInt(headerRow) - 1;

    //var headerColumns = [];
    var sheetA1Notation = '';
    for (var counter = headerRow; counter < sheetData.length; counter++) {
        if (counter == -1)//If there is no header column.
        {
            //console.log("sheet data length" + sheetData.length);
            var rowObject = getNonEmptyRowFromSheetData(sheetData);
            var sheetA1NotationWithNoHeader = [];
            if (rowObject != null && rowObject.rowData != null) {
                var nonEmptyRow = rowObject.rowData;
                for (var index = 0; index < nonEmptyRow.length; index++) {
                    sheetA1NotationWithNoHeader.push("R" + rowObject.rowNumber + "C" + (parseInt(index) + 1));
                }//end of for loop.
                sheetA1Notation = sheetA1NotationWithNoHeader;
            }
        }
        else {
            var rowData = sheetData[counter];
            if (rowData != null) {
                rowData = rowData.join('');
                //if (rowData != "") {
                if (counter == parseInt(headerRow)) {
                    sheetA1Notation = '';
                    if (sheetA1NotationDetails.length > counter) {
                        sheetA1Notation = sheetA1NotationDetails[counter];
                    }
                    //dataSourceColumns[0] = removeEmptyColumns(sheetData[counter], sheetA1Notation); //Adding header row.
                    //dataSourceColumnsAll[0] = sheetData[counter];
                }
                else {
                    sheetRecords.push(sheetData[counter]);
                }
                //}
            }
        }
    }//end of for loop.
}

function runInsertChartImageIntoSheetMethod() {
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    $(".se-pre-con").fadeIn("slow");
    logUserActionIntoDatabase(clickedChartName + "-Chart_Image_Inserted", "Charts");

    // Get this chart 
    google.script.run.withSuccessHandler(function (chartMetaObject) {
        var myChartMeta = JSON.parse(chartMetaObject);

        var chartAddedIntoMyChartListDate = new Date(+createdon); // createdOn
        selectedDimensions = JSON.parse(JSON.parse(myChartMeta.ChartMetaJSON).selectedDimensions);
        selectedMeasures = JSON.parse(JSON.parse(myChartMeta.ChartMetaJSON).selectedMeasures);
        tableHeaderRow = JSON.parse(JSON.parse(myChartMeta.ChartMetaJSON).headerRow);
        selectedChartNameFromSelectChartUI = myChartMeta.ChartName;

        // TODO, need to write its conversion code
        if (chartAddedIntoMyChartListDate < googleSheetV3LaunchDate) {
            myChartMeta.ChartDataJSON = JSON.parse(myChartMeta.ChartDataJSON);//getSampleDataRows(clickedChartName, myChartMeta.ChartDataJSON); 
        }
        else {
            myChartMeta.ChartDataJSON = convertDataIntoChartFormat(JSON.parse(myChartMeta.ChartDataJSON));
        }
        //console.log(JSON.stringify(myChartMeta.ChartDataJSON));
        drawChartForImage(myChartMeta);
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).getSelectedChartMeta(chartGuid);
}

function hideMessageDialog() {
    $(".error-overlay").fadeOut("slow");
    $(".error_dialog").fadeOut("slow");
}

function removeChartFromMyChartsLis(chartGuid) {
    // var r = confirm("Are you sure to remove it from My Charts!");
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    // if (r == true) {
    $(".se-pre-con").fadeIn("slow");
    //console.log("guid to remove = > " + chartGuid);
    google.script.run.withSuccessHandler(function (addonInitialStateObj) {

        $(".se-pre-con").fadeOut("slow");

        if (JSON.parse(addonInitialStateObj) == "RemovedSuccessfully") {
            //alert("Chart removed from My Charts list successfully!");

            // if removeChartFromMyChartsLis() method called from chart detailed view then show my charts container otherwise simply remove thumbnail
            if (removeChartFromMyChartsListClickedFromMenuItem) {
                openMyChartsContainerView();
                removeChartFromMyChartsListClickedFromMenuItem = false;
            }
            else {
                // removed relevant thumbnail
                $("#thumbnailChartGuid_" + chartGuid).remove();
            }

            showMessageDialog("Chart Removed", "Chart has been removed successfully from My Chart list.", "information", false, [], true);
            openMyChartsContainerView();//Now Added
        }
        else {
            //alert("Error while removing chart from My Charts List");
            showMessageDialog("Removal Error", "Chart is not removed from My Chart list. Please try again!", "error", false, [], true);
        }
    })
        .withFailureHandler(
        function (msg, element) {
            handleError(msg);
        }
        ).removeSelectedChartMeta(chartGuid);
    //}
}

function createGuid() {
    function S4() {
        return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
    }
    //return (S4() + S4() + "-" + S4() + "-4" + S4().substr(0, 3) + "-" + S4() + "-" + S4() + S4() + S4()).toLowerCase();
    return (S4() + S4() + "-" + S4() + "-4" + S4().substr(0, 3) + "-" + S4() + S4() + S4()).toLowerCase();
}

function loadDimensionMeasuresInEditMode() {
    //loading dimensions and measures.
    var columnsCount = 0;
    //var headerRow = localStorage.getItem("headerRow");
    var headerRow = getDateFromTempStorage("headerRow"); // contains saved header rows
    var selectedDimensions = getDateFromTempStorage("selectedDimensions"); // contains selected dimensions
    var selectedMeasures = getDateFromTempStorage("selectedMeasures"); // contains selected measures

    //console.log('di:' + selectedDimensions);
    //console.log('me:' + selectedMeasures);

    headerRow = JSON.parse(headerRow);

    if (selectedDimensions != null && selectedDimensions != "") {
        selectedDimensions = JSON.parse(selectedDimensions);
    }

    if (selectedMeasures != null && selectedMeasures != "") {
        selectedMeasures = JSON.parse(selectedMeasures);
    }

    if (headerRow.length > 0)
        columnsCount = headerRow[0].length;
    var chartObject = getSelectedChartRulesObject();
    if (chartObject != null) {
        if (selectedDimensions != null) {
            for (var counter = 0; counter < selectedDimensions.length; counter++) {
                var isSettingsColumnExistInSheet = CheckSettingsColumnExistInSheetColumns(selectedDimensions[counter]);
                addDimensionsDropDownList(selectedDimensions[counter], isSettingsColumnExistInSheet);
            }//end of for loop.
        }

        if (selectedMeasures != null) {
            for (var counter = 0; counter < selectedMeasures.length; counter++) {
                var isSettingsColumnExistInSheet = CheckSettingsColumnExistInSheetColumns(selectedMeasures[counter]);
                addMeasuresDropDownList(selectedMeasures[counter], isSettingsColumnExistInSheet);
            }
        }

    }//end of inner if statement.

    //var dataRows = localStorage.getItem("dataRows");
    var dataRows = getDateFromTempStorage("dataRows");
    dataRows = JSON.parse(dataRows);
    //console.log(dataRows.length);
    var rowsLength = 0;
    if (dataRows != null && dataRows.length > 0) {
        rowsLength = dataRows.length;
    }
    //setting range control max value.            

    $('.slidecontainer').css('display', 'block');
    $('#selectedRowsDiv').css('display', 'block');
    $('#spanSelectRowRange').css('display', 'block');
    $("#rangeValue").html(rowsLength);
    var startRange = parseInt($("#txtBoxHeaderRow").val());
    if (startRange < 1) {
        startRange = 0;
    }
    setSelectRowRangeTextboxValues(startRange, rowsLength);
    //$("#maxRangeValue").html(rowsLength);
    //}//end of outer if statement.
    // set in default as all data rows

    //updateSliderValue(0, rowsLength);
    var startRowNumber = getDateFromTempStorage("dataRowFrom");
    var endRowNumber = getDateFromTempStorage("dataRowTo");


    //console.log('start row number:' + startRowNumber);
    //console.log('end row number:' + endRowNumber);

    if (sheetRecords.length < startRowNumber) {
        startRowNumber = 1;
    }
    startRowNumber = startRowNumber + startRange;
    endRowNumber = endRowNumber + startRange;
    updateSliderValue(startRowNumber, endRowNumber);

    $('.dropdownchangeclass').removeAttr("disabled");

    var dimensionsInputLength = $('.dimension').length;
    var allowedDimensions = getSelectedChartAllowedDimensions();

    if (dimensionsInputLength >= allowedDimensions) {
        $("#addDimensionClick").hide();
    }

    var meausresInputLength = $('.metric').length;
    var allowedMetrics = getSelectedChartAllowedMeasures();

    if (meausresInputLength >= allowedMetrics) {
        $("#addMeasureClick").hide();
    }

    if (chartRequiredNoOfDimAndMetricsChosen()) {
        $("#divDrawChart").addClass('activeDrawButton');
        $('#divDrawChart').css("color", "#F37A2D");
        $('#divDrawChart').css("background-color", "white");
        $("#startRowTextBox").prop("readonly", false);
        $("#endRowTextBox").prop("readonly", false);
    }
    else {
        $("#divDrawChart").removeClass('activeDrawButton');
        $('#divDrawChart').css("color", "#B8B8B8");
        $("#startRowTextBox").prop("readonly", true);
        $("#endRowTextBox").prop("readonly", true);
    }

    bindDivsDropEvent();
}

function loadSheetsInEditMode(editMyChartObject) {
    $('#dropdownSheets').empty();
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    //$('#dropdownSheets').append($('<option value="' + newlyAddedSheetNameOnEditMode + '">' + newlyAddedSheetNameOnEditMode + '</option>'));
    //$('#dropdownSheets').append($('<option value="' + editMyChartObject.sheetName + '">' + editMyChartObject.sheetName + '</option>'));
    google.script.run.withSuccessHandler(function (sheetsName) {
        if (sheetsName != null && sheetsName != undefined) {
            //populate dropdownlist with sheetsName.
            for (var counter = 0; counter < sheetsName.length; counter++) {
                $('#dropdownSheets').append($('<option value="' + sheetsName[counter] + '">' + sheetsName[counter] + '</option>'));
            } //end of for loop. 
            //alert("editMyChartObject.sheetName " + editMyChartObject.sheetName);
            $('#dropdownSheets').val(editMyChartObject.sheetName);
            var dropdownSheets = document.getElementById('dropdownSheets');
            dropdownSheets.addEventListener("change", onSheetSelectionChange);
            //adding new sheet with data.
            //addNewSheetWithData(newlyAddedSheetNameOnEditMode, editMyChartObject);                

            loadExistingSheetInEditMode(editMyChartObject.sheetName, editMyChartObject);
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).readAllSheetsNames();
}

function loadExistingSheetInEditMode(sheetName, editMyChartObject) {
    // && editMyChartObject.headerRow != null && editMyChartObject.dataRows != null &&
    // editMyChartObject.dataRows.length > 0
    if (editMyChartObject != null) {

        changeActiveSheet(sheetName);
        //google.script.run.withSuccessHandler(function(result) {
        //console.log('result:' + result);
        //if (result == 1) {
        onSheetSelectionChange(sheetName);

        $("#txtBoxHeaderRow").val(editMyChartObject.headerRowNumber);
        //}
        //else {
        //$(".se-pre-con").fadeOut("slow");
        //}
        //})
        //.withFailureHandler(
        //function(msg, element) {
        //$(".se-pre-con").fadeOut("slow");
        //handleError(msg);
        //}
        //).addNewSheet(sheetName, JSON.stringify(editMyChartObject.headerRow), JSON.stringify(editMyChartObject.dataRows));
    }
    else {
        $(".se-pre-con").fadeOut("slow");
    }
}

var selectedChart = "";
var selectedDimensions = [];
var selectedMeasures = [];
var tableHeaderRow = [[]];
var selectedChartCategory = "general";
var windowHeight = $(this).height();
var windowWidth = $(this).width();
var Heightcharticonsmenubar = 0;//$('#charticonsmenubar').height();
var HeightSearchChartDiv = 0;//$('.SearchChartDiv').height();
var HeightOfmyChartsThumbnailsDiv = 0;//windowHeight - Heightcharticonsmenubar - HeightSearchChartDiv;
var trialSectionHeaderBarHeight = 0, selectSheetDDLContainerHeight = 0;
var mySavedListSelectedChartName, mySavedListSelectedChartCategory, mySavedListSelectedChartProperties, mySavedListSelectedChartData;
var removableMyChartName;

function setDivsHeight() {
    var defaultPaddingHeight = 10;
    var defaultDataSsourceMargin = 4;

    windowHeight = $(this).height();
    windowWidth = $(this).width();

    Heightcharticonsmenubar = $('#charticonsmenubar').outerHeight();
    HeightSearchChartDiv = $('.SearchChartDiv').outerHeight();
    trialSectionHeaderBarHeight = $(".trialSectionHeaderBar").outerHeight();
    selectSheetDDLContainerHeight = 38;//$('#selectSheetDDLContainerDiv').innerHeight();

    var createNewChartDivHeight = $('#CreateChartContentDiv').height();
    var createChartTopPadding = 10;
    trialSectionHeaderBarHeight = trialTopBarDisplayed == false ? 0 : trialSectionHeaderBarHeight;

    // in new UI v4.0, no buy now top bar so height is 0
    trialSectionHeaderBarHeight = 0;
    HeightSearchChartDiv = 0;

    HeightOfmyChartsThumbnailsDiv = windowHeight - Heightcharticonsmenubar - HeightSearchChartDiv - trialSectionHeaderBarHeight;

    $('#myChartsThumbnailsDiv').width(windowWidth);

    $('#ChartouterDiv').height(windowHeight - Heightcharticonsmenubar);//HeightOfmyChartsThumbnailsDiv + Heightcharticonsmenubar - 40);
    $('#myChartsThumbnailsDiv').css("max-height", windowHeight - Heightcharticonsmenubar - selectSheetDDLContainerHeight - createNewChartDivHeight - createChartTopPadding);
    $('#divChartExpoCharts').height(windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - HeightSearchChartDiv - defaultPaddingHeight);

    $('#divChartExpoCharts').css("overflow-y", "auto");
    $('#divChartExpoCharts').css("overflow-x", "hidden");
    var createChartButtonHeight = 37;
    //var dataSourceHeaderDivHeight = 29;//$('#DataSourceDivHeaderRow').outerHeight();
    if (!trialTopBarDisplayed) {
        //$('#DataSourceDiv').height(windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - 12);
        //$('#DataSourceDiv .containerBody').css("max-height", windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - 12 - createChartButtonHeight);
        $('#DataSourceDiv').height(windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - defaultDataSsourceMargin);
        $('#DataSourceDiv .containerBody').css("max-height", windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - defaultPaddingHeight - defaultDataSsourceMargin - createChartButtonHeight);
    }
    else {
        $('#DataSourceDiv').height(windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - defaultDataSsourceMargin);
        $('#DataSourceDiv .containerBody').css("max-height", windowHeight - trialSectionHeaderBarHeight - Heightcharticonsmenubar - defaultPaddingHeight - defaultDataSsourceMargin - createChartButtonHeight);
    }
}

function getWholeSheetRows() {
    var sheetSampleRowsCount = 50;
    var sheetSampleRows = [];

    //console.log("actualSheetData = > "+JSON.stringify(actualSheetData));

    if (sheetSampleRowsCount < actualSheetData.length) {
        //alert(actualSheetData.length);
        for (var i = 0; i < actualSheetData.length && i < sheetSampleRowsCount; i++) {
            sheetSampleRows.push(actualSheetData[i]);
        }
        return sheetSampleRows;
    }
    //alert(actualSheetData.length);
    return actualSheetData;
}

// triggeredPoint will be either from sample data or from sheet data
function getProcessedDataObjectForChartViewer(triggeredPoint) {
    var processedObject = {};
    if (triggeredPoint == "sampledata") {
        //  construct data
        // 1. get header rows
        processedObject.selectedChart = selectedChart;
        processedObject.selectedChartDisplayName = selectedchartDisplayName;
        processedObject.selectedChartCategory = selectedChartCategory;
        processedObject.headerRow = getSampleDataHeaderRow(selectedChart);
        processedObject.dataRows = getSampleDataRows(selectedChart);
        processedObject.defaultProperties = null; // Todo get default properties  

        processedObject.useHeaderRow = "true";

        // 2. get data rows
    }
    else if (triggeredPoint == "sheetdata") {
        //  from sheet data
    }

    return processedObject;
}

function storeDataInLocalStorage(storableObject) {
    if (localStorageAccessible) {
        window.localStorage.clear();
        window.localStorage.setItem('selectedChart', storableObject.selectedChart);
        window.localStorage.setItem('selectedChartDisplayName', storableObject.selectedChartDisplayName);
        window.localStorage.setItem('selectedChartCategory', storableObject.selectedChartCategory);
        window.localStorage.setItem('headerRow', JSON.stringify(storableObject.headerRow));
        window.localStorage.setItem('dataRows', JSON.stringify(storableObject.dataRows));
        window.localStorage.setItem("sameContractList", JSON.stringify(samecontractlist));
        window.localStorage.setItem("editableChartCustomName", currentEditableCustomChartName);

        window.localStorage.setItem('a1NotationInformation', storableObject.a1NotationInformation);

        window.localStorage.setItem('selectedDimensions', JSON.stringify(storableObject.selectedDimensions));
        window.localStorage.setItem('selectedMeasures', JSON.stringify(storableObject.selectedMeasures));

        window.localStorage.setItem('useHeaderRow', storableObject.useHeaderRow);
        window.localStorage.setItem("headerRowAnnotation", JSON.stringify(storableObject.headerRowAnnotation));
        window.localStorage.setItem("dataRowsAnnotation", JSON.stringify(storableObject.dataRowsAnnotation));
        window.localStorage.setItem("fileName", storableObject.fileName);
        window.localStorage.setItem("fileId", storableObject.fileId);
        window.localStorage.setItem("sheetName", storableObject.sheetName);
        window.localStorage.setItem("sheetId", storableObject.sheetId);
        window.localStorage.setItem("headerRowNumber", storableObject.headerRowNumber);
        window.localStorage.setItem("dataRowFrom", storableObject.dataRowFrom);
        window.localStorage.setItem("dataRowTo", storableObject.dataRowTo);
        window.localStorage.setItem("sheetWholedataRows", JSON.stringify(getWholeSheetRows()));

        //alert(exploreSampleDataLinkClicked);
        if (exploreSampleDataLinkClicked) {
            window.localStorage.setItem('openDataViewerOnLoad', "true");
            exploreSampleDataLinkClicked = false;
        }
        else {
            window.localStorage.setItem('openDataViewerOnLoad', "false");
        }

        if (storableObject.myCharts != undefined && storableObject.myCharts == "true") {
            window.localStorage.setItem('myChart', storableObject.myCharts);
            window.localStorage.setItem('dimension', JSON.stringify(storableObject.dimension));
            window.localStorage.setItem('defaultProperties', JSON.stringify(storableObject.defaultProperties));
            window.localStorage.setItem('editableChartGuid', storableObject.editableChartGuid);
        }
        else {
            window.localStorage.setItem('myChart', "false");
        }
    }
    else {
        storableObjectInTempStorage['selectedChart'] = storableObject.selectedChart;
        storableObjectInTempStorage['selectedChartDisplayName'] = storableObject.selectedChartDisplayName;
        storableObjectInTempStorage['selectedChartDisplayName'] = storableObject.selectedChartDisplayName;
        storableObjectInTempStorage['selectedChartCategory'] = storableObject.selectedChartCategory;
        storableObjectInTempStorage['headerRow'] = JSON.stringify(storableObject.headerRow);
        storableObjectInTempStorage['dataRows'] = JSON.stringify(storableObject.dataRows);
        storableObjectInTempStorage['sameContractList'] = JSON.stringify(samecontractlist);
        storableObjectInTempStorage['editableChartCustomName'] = currentEditableCustomChartName;

        storableObjectInTempStorage['selectedDimensions'] = JSON.stringify(storableObject.selectedDimensions);
        storableObjectInTempStorage['selectedMeasures'] = JSON.stringify(storableObject.selectedMeasures);

        storableObjectInTempStorage['a1NotationInformation'] = storableObject.a1NotationInformation;
        storableObjectInTempStorage['useHeaderRow'] = storableObject.useHeaderRow;
        storableObjectInTempStorage['headerRowAnnotation'] = JSON.stringify(storableObject.headerRowAnnotation);
        storableObjectInTempStorage['dataRowsAnnotation'] = JSON.stringify(storableObject.dataRowsAnnotation);
        storableObjectInTempStorage['fileName'] = storableObject.fileName;
        storableObjectInTempStorage['fileId'] = storableObject.fileId;
        storableObjectInTempStorage['sheetName'] = storableObject.sheetName;
        storableObjectInTempStorage['sheetId'] = storableObject.sheetId;
        storableObjectInTempStorage['headerRowNumber'] = storableObject.headerRowNumber;
        storableObjectInTempStorage['dataRowFrom'] = storableObject.dataRowFrom;
        storableObjectInTempStorage['dataRowTo'] = storableObject.dataRowTo;
        storableObjectInTempStorage['sheetWholedataRows'] = JSON.stringify(getWholeSheetRows());
        storableObjectInTempStorage['trialExpired'] = storableObject.trialExpired; // flag to be passed for ChartViewer and show water mark accordingly

        storableObjectInTempStorage['packageStatus'] = storableObject.packageStatus;
        storableObjectInTempStorage['loggedInUserType'] = storableObject.loggedInUserType;


        if (exploreSampleDataLinkClicked) {
            storableObjectInTempStorage['openDataViewerOnLoad'] = "true";
            exploreSampleDataLinkClicked = false;
        }
        else {
            storableObjectInTempStorage['openDataViewerOnLoad'] = "false";
        }

        //console.log("storableObject.myCharts = " + storableObject.myCharts + " storableObject.dimension=> " + storableObject.dimension + " synchedChartDimensions=> " + JSON.stringify(synchedChartDimensions));

        if (storableObject.myCharts != undefined && storableObject.myCharts == "true") {

            if (chartInEditModeSyncTime) {
                //alert("if = synchedChartGUID " + synchedChartGUID);
                storableObjectInTempStorage['dimension'] = JSON.stringify(synchedChartDimensions);//storableObject.dimension;
                storableObjectInTempStorage['defaultProperties'] = storableObject.defaultProperties;
                storableObjectInTempStorage['editableChartGuid'] = synchedChartGUID;
                storableObjectInTempStorage['myChart'] = storableObject.myCharts; // this variable used in chartviewer to treat it chart opened from my chart
            }
            else {
                // alert("eseif");
                storableObjectInTempStorage['myChart'] = storableObject.myCharts;
                storableObjectInTempStorage['dimension'] = JSON.stringify(storableObject.dimension);
                storableObjectInTempStorage['defaultProperties'] = JSON.stringify(storableObject.defaultProperties);
                storableObjectInTempStorage['editableChartGuid'] = storableObject.editableChartGuid;

                // processedObject.editableChartGuid = synchedChartGUID;
                if (storableObjectInTempStorage['dimension'] == undefined) {
                    storableObjectInTempStorage['dimension'] = JSON.stringify(synchedChartDimensions);
                    storableObjectInTempStorage['defaultProperties'] = JSON.stringify(synchedChartProperties);
                }

                /*
                console.log("storableObjectInTempStorage['defaultProperties']=>");
                console.log(JSON.stringify(storableObjectInTempStorage['defaultProperties']));
                console.log("storableObjectInTempStorage['dimension']=>");
                console.log(JSON.stringify(storableObjectInTempStorage['dimension']));
                */
            }
        }
        else {
            storableObjectInTempStorage['myChart'] = "false";
        }

        // console.log("at create chart storableObjectInTempStorage['dimension'] = " + JSON.stringify(storableObjectInTempStorage['dimension']));
        // console.log("at create chart storableObjectInTempStorage['defaultProperties'] = " + JSON.stringify(storableObjectInTempStorage['defaultProperties']));
    }
}

function initializeChartViewer(triggeredPoint, headerRow, dataRows, selectedChart, selectedChartDisplayName) {
    // do not disabled it, as latest state retrieved at server end and updated as well
    //clearTimeout(syncSheetDataWithAddonTimerHandler); // disable sync, if needs to sync, then first get latest state from server like its properties then update with latest.
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    if (triggeredPoint.toLowerCase() == "sampledata" || triggeredPoint.toLowerCase() == "mycharts" || triggeredPoint.toLowerCase() == "editdata") {
        currentEditableCustomChartName = '';
        google.script.run.showChartViewerInDialog();
    }
    else {
        google.script.run.withSuccessHandler(function () {
            var sheetDetailObject = google.script.run.getFileSheetDetail();
            if (sheetDetailObject != null) {
                sheetName = sheetDetailObject.sheetName;
                sheetId = sheetDetailObject.sheetId;
                fileName = sheetDetailObject.fileName;
                fileId = sheetDetailObject.fileId;
            }
            var processedObject = {};
            processedObject.selectedChart = selectedChart;
            processedObject.selectedChartDisplayName = selectedchartDisplayName;
            processedObject.selectedChartCategory = selectedChartCategory;
            processedObject.headerRow = headerRow;
            processedObject.dataRows = dataRows;
            processedObject.defaultProperties = null; // Todo get default properties

            processedObject.a1NotationInformation = oldSelectedCellsA1Notations;
            processedObject.useHeaderRow = $('#chkHeaderRow').prop("checked");
            processedObject.sheetName = sheetName;
            processedObject.sheetId = sheetId;
            processedObject.fileName = fileName;
            processedObject.fileId = fileId;
            processedObject.headerRowNumber = $("#txtBoxHeaderRow").val();
            processedObject.dataRowFrom = +$("#startRowTextBox").val() - 2; //minValue;
            processedObject.dataRowTo = +$("#endRowTextBox").val() - 2; //maxValue;
            processedObject.headerRowAnnotation = headerRowAnnotation;
            processedObject.dataRowsAnnotation = dataRowsAnnotation;

            //console.log(selectedChart + " " + selectedchartDisplayName + " " + selectedChartCategory + " " + JSON.stringify(headerRow) + " " + JSON.stringify(dataRows));
            //console.log(sheetName + " " + fileName + " " + minValue + " " + maxValue + " " + JSON.stringify(headerRowAnnotation) + " " + JSON.stringify(dataRowsAnnotation));
            storeDataInLocalStorage(processedObject);

            google.script.run.showChartViewerInDialog();
        })
            .withFailureHandler(
            function (msg, element) {
                handleError(msg);
            }
            ).getFileSheetDetail();
    }

    setTimeout(function () {
        //$("#divDrawChart").css('cursor', 'pointer');
        $(".se-pre-con").fadeOut("slow");
    }, 5000);
}

// 1. Saved chart state into database on server
// 2. After successfull saving, open chart viewer
// 3. 
function openChartViewer(triggeredPoint, headerRow, dataRows, selectedChart, selectedChartDisplayName) {
    //alert("openChartViewer");
    //  1. get processed data
    // 2. store it inside of local storage
    // 3. open chart viewer dialog
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    if (localStorageAccessible) {
        initializeChartViewer(triggeredPoint, headerRow, dataRows, selectedChart, selectedChartDisplayName);
    }
    else {
        //console.log(JSON.stringify(storableObjectInTempStorage));
        if (triggeredPoint.toLowerCase() == "sampledata" || triggeredPoint.toLowerCase() == "mycharts" || triggeredPoint.toLowerCase() == "editdata") {
            /* currentEditableCustomChartName = '';
         }
         else {*/
            currentEditableCustomChartName = '';
            var processedObject = {};
            // alert("synchedChartGUID = " + synchedChartGUID);
            // console.log("synchedChartGUID = " + synchedChartGUID);
            if (triggeredPoint.toLowerCase() == "mycharts" || triggeredPoint.toLowerCase() == "editdata") {
                processedObject.myCharts = "true";
                processedObject.editableChartGuid = synchedChartGUID;
            }

            //alert("openChartViewer processedObject.myChart=" + processedObject.myChart + " triggeredPoint.toLowerCase()= " + triggeredPoint.toLowerCase());

            processedObject.selectedChart = selectedChart;
            processedObject.selectedChartDisplayName = selectedchartDisplayName;
            processedObject.selectedChartCategory = selectedChartCategory;
            processedObject.headerRow = headerRow;



            processedObject.dataRows = dataRows;
            processedObject.defaultProperties = synchedChartProperties; // Todo get default properties

            processedObject.a1NotationInformation = oldSelectedCellsA1Notations;
            processedObject.useHeaderRow = $('#chkHeaderRow').prop("checked");
            processedObject.sheetName = sheetName;
            processedObject.sheetId = '';
            processedObject.fileName = fileName;
            processedObject.fileId = '';
            processedObject.headerRowNumber = $("#txtBoxHeaderRow").val();
            processedObject.dataRowFrom = +$("#startRowTextBox").val() - 2; //minValue;
            processedObject.dataRowTo = +$("#endRowTextBox").val() - 2; //maxValue;
            processedObject.headerRowAnnotation = headerRowAnnotation;
            processedObject.dataRowsAnnotation = dataRowsAnnotation;
            processedObject.trialExpired = trialExpired; // flag to be passed for ChartViewer and show water mark accordingly
            processedObject.packageStatus = licenceMessage; // flag contains user package status
            processedObject.loggedInUserType = loggedInUserType; // logged in user type, can be [normal user, domain admin or domain normal user]

            if (storableObjectInTempStorage["selectedDimensions"] != null && storableObjectInTempStorage["selectedDimensions"] != "")
                processedObject.selectedDimensions = JSON.parse(storableObjectInTempStorage["selectedDimensions"]);

            if (storableObjectInTempStorage["selectedMeasures"] != null && storableObjectInTempStorage["selectedMeasures"] != "")
                processedObject.selectedMeasures = JSON.parse(storableObjectInTempStorage["selectedMeasures"]);


            storeDataInLocalStorage(processedObject);
        }

        //console.log(JSON.stringify(dataRows));

        //alert("storableObjectInTempStorage.editableChartGuid on saving time=> "+storableObjectInTempStorage.editableChartGuid);
        // upload temp data on server and then send call for open chartViewer
        var saveDataInTempStorage = {};
        saveDataInTempStorage.RecentStoredObject = JSON.stringify(storableObjectInTempStorage);
        saveDataInTempStorage.IP = "192.168.192.8"; // get user system IP in javascript
        // before saving replace <br> with __br__ as backend web api not saved <br>
        saveDataInTempStorage = JSON.stringify(saveDataInTempStorage);
        var newsaveDataInTempStorage = saveDataInTempStorage.replace("<br>", "___br___");

        google.script.run.withSuccessHandler(function () {
            initializeChartViewer(triggeredPoint, headerRow, dataRows, selectedChart, selectedChartDisplayName);
        })
            .withFailureHandler(
            function (msg, element) {
                //$(".se-pre-con").fadeOut("slow");
                handleError(msg);
            }
            ).saveDataInTempStorage(newsaveDataInTempStorage);

    }
}

// Get sample Data
function getData() {
    try {
        //set dataSource detail as well
        dataSourceType = "sampleData";
        dataSourceRange = "SampleData." + selectedChart;
        dataSourceSheetName = "";
        //alert(selectedChartCategory);

        if (selectedChartCategory !== undefined && selectedChartCategory === "PPC") {
            return PPCChartsSampleData[selectedChart]();
        }
        else {
            return SampleData[selectedChart]();
        }
    }
    catch (ex) {
        LogError(ex.message);
    }
}

function getSampleDataHeaderRow(selectedChart, maxNoOfColumnsAsPerSavedData) {
    try {
        var dataHeaders = columnNameMapper[selectedChart];
        var headerRowColumns = [];
        var headerRow = [];

        if (selectedChart == "SentimentTrendChart") {
            //dataHeaders = sentimentTrendGetDataHeaderForExcel(data);
            //  convert object data into a 2d array
            for (var j = 0; j < dataHeaders.length; j++) {
                headerRowColumns.push(dataHeaders[j]);
            }
        }
        else {
            // Below couple of lines added to handle all my saved charts whose header is not saved
            // so we find them by self. it can be either more or less than saved headers in mapper file
            if (maxNoOfColumnsAsPerSavedData != undefined) {
                if (dataHeaders > maxNoOfColumnsAsPerSavedData) {
                    var j = 0;
                    for (var j = 0; j < maxNoOfColumnsAsPerSavedData - 1; j++) {
                        headerRowColumns.push(dataHeaders[j]);
                    }
                    headerRowColumns.push(dataHeaders[j + 1]);
                }
                else {
                    // if dataHeaders columns lesser than saved data then need to add additional 
                    for (var j = 0; j < maxNoOfColumnsAsPerSavedData - 1; j++) {
                        headerRowColumns.push("Topic" + j);
                    }
                    // at end add count column
                    headerRowColumns.push("Count");
                }
            }
            else {
                // convert object data into a 2d array
                for (var j = 0; j < dataHeaders.length; j++) {
                    headerRowColumns.push(dataHeaders[j]);
                }
            }
        }

        headerRow = [headerRowColumns];

        return headerRow;
        // google.script.run.insertData(JSON.stringify(headerRow), JSON.stringify(dataRows));
    }
    catch (ex) {
        LogError(ex.message);
    }
    return [[]];
}

function getSampleDataRows(selectedChart, inputDataInCaseOfMyChartSavedData) {
    try {
        var data = getData();

        if (inputDataInCaseOfMyChartSavedData != undefined) {
            data = inputDataInCaseOfMyChartSavedData;
        }

        var dataRows = [];

        if (selectedChart == "SentimentTrendChart") {
            dataRows = sentimentTrendGetDataRowForExcel(data);
        }
        else {
            // for multi measure charts call MultiMeasure method
            if (selectedChart == "GaugeChart" || selectedChart == "ISGraph" || selectedChart == "ScatterChartAdvance" ||
                selectedChart == "HierarichalBarChartAdvance" || selectedChart == "ChordChart" || selectedChart == "DoubleMeasureComparisonChart") {
                if (selectedChart == "GaugeChart") {
                    var gaugeData = [];
                    gaugeData.push(data)
                    data = gaugeData;
                }
                dataRows = getDataRowForExcelMultiMeasure(data);
            }
            else {
                dataRows = getDataRowForExcel(data);
            }
        }

        return dataRows;
    }
    catch (ex) {
        LogError(ex.message);
    }
    return [[]];
}

function insertSampleData(selectedChart) {
    try {
        //logUserActionIntoDatabase(selectedChart + "-DataInsertedIntoSheet", "Charts");
        var data = getData();
        //data = selectedChartData;
        var dataHeaders = columnNameMapper[selectedChart];
        var headerRowColumns = [];
        var headerRow = [];
        var dataRows = [];

        if (selectedChart == "SentimentTrendChart") {
            //dataHeaders = sentimentTrendGetDataHeaderForExcel(data);
            // convert object data into a 2d array
            for (var j = 0; j < dataHeaders.length; j++) {
                headerRowColumns.push(dataHeaders[j]);
            }
            dataRows = sentimentTrendGetDataRowForExcel(data);
        }
        else {
            // convert object data into a 2d array
            for (var j = 0; j < dataHeaders.length; j++) {
                headerRowColumns.push(dataHeaders[j]);
            }

            // for multi measure charts call MultiMeasure method
            if (selectedChart == "ISGraph" || selectedChart == "ScatterChartAdvance" ||
                selectedChart == "HierarichalBarChartAdvance" || selectedChart == "ChordChart" || selectedChart == "DoubleMeasureComparisonChart") {
                dataRows = getDataRowForExcelMultiMeasure(data);
            }
            else {
                dataRows = getDataRowForExcel(data);
            }
        }

        headerRow = [headerRowColumns];

        var sampleDataForExcel = {};// Todo //new Office.TableData( dataRows, dataHeaders);

        //google.script.run.insertData(JSON.stringify(headerRow), JSON.stringify(dataRows));
    }
    catch (ex) {
        LogError(ex.message);
    }
}

function convertDataIntoExcelFormat(data) {
    var convertedData = [];
    var subCategories = [];
    var row = JSON.parse(JSON.stringify(tabularDataCharts_DimensionColumnName[selectedChart]));
    var found = false;
    if (tabularDataCharts_DimensionColumnName[selectedChart].length == 1) {
        for (var i = 0; i < data.length; i++) {
            for (var j = 0; j < data[i].subCategory.length; j++) {
                if (subCategories.indexOf(data[i].subCategory[j].name) == -1) {
                    subCategories.push(data[i].subCategory[j].name);
                }
            }
        }
        for (var i = 0; i < subCategories.length; i++) {
            row.push(subCategories[i]);
        }
        convertedData.push(row);
        for (var i = 0; i < data.length; i++) {
            row = [];
            row.push(data[i].category || data[i].name);
            for (var k = 0; k < subCategories.length; k++) {
                found = false;
                for (var j = 0; j < data[i].subCategory.length; j++) {
                    if (subCategories[k] == data[i].subCategory[j].name) {
                        row.push(+data[i].subCategory[j].val);
                        found = true;
                        break;
                    }

                }
                if (found == false) {
                    row.push(0);
                }
            }
            convertedData.push(row);
        }
    }
    else if (tabularDataCharts_DimensionColumnName[selectedChart].length == 2) {
        for (var i = 0; i < data.length; i++) {
            for (var j = 0; j < data[i].subCategory.length; j++) {
                for (var k = 0; k < data[i].subCategory[j].subCategory.length; k++) {
                    if (subCategories.indexOf(data[i].subCategory[j].subCategory[k].name) == -1) {
                        subCategories.push(data[i].subCategory[j].subCategory[k].name);
                    }
                }
            }
        }
        for (var i = 0; i < subCategories.length; i++) {
            row.push(subCategories[i]);
        }
        convertedData.push(row);
        for (var i = 0; i < data.length; i++) {
            for (var n = 0; n < data[i].subCategory.length; n++) {
                row = [];
                row.push(data[i].category || data[i].name);
                row.push(data[i].subCategory[n].category || data[i].subCategory[n].name);
                for (var k = 0; k < subCategories.length; k++) {
                    found = false;
                    for (var j = 0; j < data[i].subCategory[n].subCategory.length; j++) {
                        if (subCategories[k] == data[i].subCategory[n].subCategory[j].name) {
                            row.push(+data[i].subCategory[n].subCategory[j].val);
                            found = true;
                            break;
                        }
                    }
                    if (found == false) {
                        row.push(0);
                    }
                }
                convertedData.push(row);
            }
        }
    }
    return convertedData;
}

function getDataRowForExcel(data) {
    try {
        var rows = [];

        function constructDataRow(data, names) {
            names = names || [];
            data.forEach(function (a) {
                if (a.hasOwnProperty("lable")) {
                    var n = names.concat(a.lable);
                }
                else if (a.hasOwnProperty("name")) {
                    var n = names.concat(a.name);
                }
                else if (a.hasOwnProperty("category")) {
                    var n = names.concat(a.category);
                }

                if (Array.isArray(a.subCategory)) {
                    constructDataRow(a.subCategory, n);
                }
                else {
                    if (a.hasOwnProperty("val")) {
                        n = n.concat((a.val.toString()));
                    }

                    rows.push(n);
                }
            });
        }
        constructDataRow(data);
        return rows;
    }
    catch (ex) {
        LogError(ex.message);
    }
}

function getDataRowForExcelMultiMeasure(data) {
    try {
        var rows = [];

        function constructDataRow(data, names) {
            names = names || [];
            data.forEach(function (a) {
                if (a.hasOwnProperty("lable")) {
                    var n = names.concat(a.lable);
                }
                else if (a.hasOwnProperty("name")) {
                    var n = names.concat(a.name);
                }
                else if (a.hasOwnProperty("category")) {
                    var n = names.concat(a.category);
                }

                if (Array.isArray(a.subCategory)) {
                    constructDataRow(a.subCategory, n);
                }
                else {
                    if (a.hasOwnProperty("val")) {
                        n = n.concat((a.val.toString()));
                    }
                    if (a.hasOwnProperty("val1")) {
                        n = n.concat((a.val1.toString()));
                    }

                    if (selectedChart == "HierarichalBarChartAdvance" || selectedChart == "ChordChart") {
                        if (a.hasOwnProperty("val2")) {
                            n = n.concat((a.val2.toString()));
                        }
                    }

                    rows.push(n);
                }
            });
        }
        constructDataRow(data);
        return rows;
    }
    catch (ex) {
        LogError(ex.message);
    }
}

function LogError(message) {
    alert("LogError => " + message);
    //debugger;
    //alert("LogError -> "+message);
    google.script.run.logExceptionInConsole(message);
}

function sentimentTrendGetDataRowForExcel(data) {
    try {
        var rows = [];
        data.forEach(function (a) {
            var row = [];
            row.push(a.lable);
            //row.push(a.val);
            row.push(a.positive);
            row.push(a.negative); //  TODO: set absolute values in sample data, currently multiply it with -1 to get absolute values
            //row.push(a.delta);

            rows.push(row);
        });
        return rows;
    }
    catch (ex) {
        LogError(ex.message);
    }
}

function paretoGroupedChartGetDataRowForExcel(data) {
    try {
        var rows = [];
        data.forEach(function (a) {
            var row = [];
            row.push(a.lable);
            row.push(a.val);
            row.push(a.previous);

            rows.push(row);
        });
        return rows;
    }
    catch (ex) {
        LogError(ex.message);
    }
}
//openChartViewer("sampledata");

function showMessageDialog(title, message, messageType, showButtons = false, buttonTypesArray, autoHide = true) {
    $(".error_title").html(title);
    $(".error_message").html(message);
    $("#anchorGetHelp").hide();
    $(".error_dialog_button_view").css("display", "block");

    if (title == "Not a valid user") {
        var learnMoreLink = '<a href = "https://chartexpo.com/home/addonhelp#googlesheetFAQ" target="_blank">Learn more</a>';
        $(".error_message").html(message + " " + learnMoreLink);
    }

    // messageType can be of following types: Error, Information,  Confirmation
    // Error => title and icon will be red
    // Information => title and icon will be green
    // Confirmation => title and icon will be orange

    if (showButtons && title == "Not a valid user") {
        $(".error_dialog_button_view").css("display", "none");
        $(".error_dialog_buttons").css("display", "block");
        $("#anchorGetHelp").show();
        $(".error_dialog_button_cancel").val(buttonTypesArray[0]);
    }
    else if (showButtons) {
        $(".error_dialog_buttons").css("display", "block");

        $(".error_dialog_button_view").val(buttonTypesArray[0]);
        $(".error_dialog_button_cancel").val(buttonTypesArray[1]);
    }
    else {
        $(".error_dialog_buttons").css("display", "none");
    }

    if (messageType === "error") {
        $(".error_title").css("color", "#ED1C24");
        $(".error_dialog_image").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/errordialog/delete.png");
    }
    else if (messageType === "information") {
        $(".error_title").css("color", "#0A9446");
        $(".error_dialog_image").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/errordialog/success.png");
    }
    else if (messageType === "confirmation") {
        $(".error_title").css("color", "#FCB040");
        $(".error_dialog_image").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/errordialog/warning.png");
    }
    $(".error_dialog_close_icon").css("display", "block");

    $(".error-overlay").fadeIn("slow");
    $(".error_dialog").fadeIn("slow");

    if (autoHide) {
        autoHideMessageDialog();
    }
}

function showMultipleAccountsLoginDialog(){
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    else {
        $(".error-overlay").fadeIn("slow");
        $(".multipleLoginMessageContainerDiv").fadeIn("slow");
    }
}

function goToHomePage() {
    if (trialExpired) {
        return;
        //showTrialTopBar(true, licenceMessage, -1, undefined, undefined, undefined);
        //ShowMainMenuViews('licenseKeyContainer');
        //viewScreen("SubscriptionScreen");
    }
    else {
        userClickedOnClearAndFlushButton = false;
        $("#licenseKeyContainer").css("display", "none");
        $("#editLicenseKeyContainer").css("display", "none");
        $('#subscriptionScreensHeaderDiv').css("display", "none");
        $("#screenManageTeamUsersForTrialPackage").css("display", "none");

        $("#subscriptionPriceRangeScreen").css("display", "none");
        $("#priceRangeDetailedScreen").css("display", "none");
        $("#subscriptionSettingsContainer").css("display", "none");
        $("#chartListDiv").css("display", "none");
        $("#myChartsContainerDiv").css("display", "none");
        $("#mainDropDownContainer").css("display", "none");
        $("#propSettingsDivgContiner").css("display", "none");
        //viewScreen("");
        $('#ChartouterDiv').show();
        boldSelectedTopMenuOption("Bold", "normal", "normal", "normal", "normal", "normal");
        openMyChartsContainerView();
    }
}

var chartObj;

function drawChartForImage(mySavedChartMeta) {
    var selectedChart = mySavedChartMeta.ChartName;
    var savedSettings = mySavedChartMeta.ChartMetaJSON;
    var chartDimension = {}, mySavedListSelectedChartCategory, mySavedListSelectedChartProperties,
        mySavedListSelectedChartData;


    if (savedSettings != undefined) {

        $("#renderToDivWithCSS").css("display", "block");

        savedSettings = JSON.parse(savedSettings);
        mySavedListSelectedChartCategory = savedSettings.chartCategory;
        mySavedListSelectedChartProperties = savedSettings.props;
        chartDimension = savedSettings.dimension;

        if (mySavedListSelectedChartProperties != undefined) {
            mySavedListSelectedChartProperties = JSON.stringify(mySavedListSelectedChartProperties).replace("___br___", "<br>");
        }

        mySavedListSelectedChartData = mySavedChartMeta.ChartDataJSON;

        var tooltipData = {
            SmartLinking: [
                { name: "SmartLinking 1", val: "customInfo" },
                { name: "SmartLinking 2", val: "customInfo" },
                { name: "SmartLinking 3", val: "customInfo" },
                { name: "SmartLinking 4", val: "customInfo" }
            ],
            Drilling: [
                { name: "Drilling 1", val: "customInfo" },
                { name: "Drilling 2", val: "customInfo" },
                { name: "Drilling 3", val: "customInfo" },
                { name: "Drilling 4", val: "customInfo" }
            ]
        }
        var settings = {
            height: chartDimension.height,
            width: chartDimension.width,
            renderTo: "renderToDiv",
            data: mySavedListSelectedChartData,
            lib: "D3",
            tooltipData: [tooltipData]
        };

        chartObj = new ChartExpo[selectedChart](settings, JSON.parse(mySavedListSelectedChartProperties));
        chartObj.draw();


        // Create chart image
        createNewChartImage();


        setTimeout(function () {
            $(".se-pre-con").fadeOut("slow");
        }, 8000);
    }

    function createNewChartImage() {
        var renderToDivWidth = 20, renderToDivHeight = 20;
        //d3.select("#renderToDiv").style("border", "1px solid #999999");

        if (selectedChart == "SankeySentimentChartAdvance" || selectedChart == "SankeyNonSentimentChartAdvance") {
            // with sort icon, chart image not exported
            d3.selectAll(".sankeySortIcon").remove();
        }

        var elmnt = document.getElementById("renderToDiv");
        if (selectedChart == "SparklineChart") {
            renderToDivHeight = +($("#renderToDiv").css("height").replace("px", "")) + 20;//elmnt.offsetHeight + 10;
            renderToDivWidth = +($("#renderToDiv").css("width").replace("px", "")) + 15;//elmnt.offsetWidth + 10;
        }
        else {
            renderToDivHeight = +($("#renderToDiv").attr("_height")) + 20;//elmnt.offsetHeight + 10;
            renderToDivWidth = +($("#renderToDiv").attr("_width")) + 15;//elmnt.offsetWidth + 10;
        }

        // set canvas size as per size of renderToDiv size
        d3.select("#canvas").attr("width", renderToDivWidth);
        d3.select("#canvas").attr("height", renderToDivHeight);

        var canvas = document.getElementById("canvas"),
            html_container = document.getElementById("renderToDivWithCSS"),
            html = html_container.innerHTML;

        //var chartDefaultPropertiesList = chartObj.getProperties(true);
        //var savedProperties = currentChartChangedPropertiesList;//getStorablePropertiesVersion(chartDefaultPropertiesList);

        // get html and generate an image in canvas
        rasterizeHTML.drawHTML(html, canvas);
        rasterizeHTML.drawHTML(html).then(function (renderResult) {
            var base64Image = canvas.toDataURL("image/png");
            google.script.run.insertChartImage(base64Image, selectedChart);//, JSON.stringify(currentChartChangedPropertiesList), JSON.stringify(selectedChartData), dataSourceType, dataSourceRange, dataSourceSheetName, (renderToDivWidth + "[]" + renderToDivHeight));

            d3.select("#canvas").style("display", "none");
            //d3.select("#renderToDiv").style("border", "0px solid #999999");

            $("#renderToDivWithCSS").css("display", "none");

        });
    }
}

function setMenuIconsWidth(totalwidth, SubscriptionDomain, isDomainNormalUser) {
    //console.log("SubscriptionDomain in setmenuicon = " + SubscriptionDomain);
    /*
    var TotalMenuWidth = totalwidth;
    var EachMenuWidth = 0;
    var removemychartspace = 0;
    var removehelpspace = 0;
    var mychartwidth = 0;
    var helpwidth = 0;
    */
    //alert(SubscriptionDomain);
    if (SubscriptionDomain == "Individual") {
        $('.managedomain').hide();
        /*
        EachMenuWidth = 300 / 3;
        removemychartspace = 15;
        removehelpspace = 15;
        mychartwidth = EachMenuWidth - removemychartspace;
        helpwidth = EachMenuWidth - removehelpspace;
        $('.managedomain').hide();
        $('.mychart').width(mychartwidth);
        $('.viewsubscription').width(EachMenuWidth + removemychartspace + removehelpspace);
        $('.help').width(helpwidth);
        */
    }
    else {
        if (!isDomainNormalUser) {
            /*
            EachMenuWidth = 300 / 4;
            removemychartspace = 15;
            removehelpspace = 15;
            mychartwidth = EachMenuWidth - removemychartspace;
            helpwidth = (EachMenuWidth - removehelpspace) / 2;
            */
            $('.managedomain').show();
            /*
            $('.mychart').width(mychartwidth);
            $('.viewsubscription').width((EachMenuWidth + helpwidth) - 14);
            $('.managedomain').width((EachMenuWidth + helpwidth) + 8);
            $('.help').width(helpwidth);*/
        }
        else {
            /*
            EachMenuWidth = 300 / 2;
            removemychartspace = 15;
            removehelpspace = 15;

            mychartwidth = EachMenuWidth;
            helpwidth = EachMenuWidth;
            */

            $('.managedomain').hide();
            $('.viewsubscription').hide();
            $('.topBarNotificationIcon').hide();
            $('#charticonsmenubar').show();

            //$('.mychart').width(mychartwidth);
            //$('.viewsubscription').width((EachMenuWidth + helpwidth) - 14);
            //$('.managedomain').width((EachMenuWidth + helpwidth) + 8);
            // $('.help').width(helpwidth);
        }
    }
    // Show top menu bars only once package detail and user type found
    $('#charticonsmenubar').show();
}
// In this method, we calculate address of cells needs to be highlighted with the help of its header row
function updateSelectedChartA1NotificationColumns(startRange, endRange) {
    //console.log(startRange);
    //console.log(endRange);
    //console.trace();
    //console.log("updateSelectedChartA1NotificationColumns called");
    var selectedChartA1NotationColumnsWithRows = [];
    //console.log(JSON.stringify(selectedChartA1NotationColumns));
    var headerRowCellsA1NotationsDetail = selectedChartA1NotationColumns;
    startSliderRange = startRange;
    endSliderRange = endRange;

    headerRowAnnotation = [];
    dataRowsAnnotation = [];

    var totalRowsToSelect = endRange - startRange;
    //console.log(startRange + "," + endRange + " data=> " + JSON.stringify(headerRowCellsA1NotationsDetail));
    if (totalRowsToSelect >= 0) {
        for (var counter = 0; counter < headerRowCellsA1NotationsDetail.length; counter++) {

            var columnHeaderCellA1Notation = headerRowCellsA1NotationsDetail[counter];

            var numericValue = columnHeaderCellA1Notation.replace(/[^0-9]+/gi, "");  // 2 - R1C2

            var rFrom = (columnHeaderCellA1Notation.split('R')[1]).split('C')[0];
            var cNumber = ((columnHeaderCellA1Notation.split('R')[1]).split('C')[1]);

            var highlightRowsFrom = parseInt(numericValue) + startRange;
            var highlightRowsUptoNumber = highlightRowsFrom + totalRowsToSelect;

            if (startRange > 0) { //its mean user has selected rows after header row, otherwise selection starts from header row
                //if header row exist then add notation for header.
                if ($("#chkHeaderRow").prop("checked")) {
                    selectedChartA1NotationColumnsWithRows.push(columnHeaderCellA1Notation);
                    headerRowAnnotation.push(columnHeaderCellA1Notation);
                }
            }

            //console.log("start range:" + startRange);
            //console.log("end range:" + endRange);
            //console.log("r from:" + rFrom);
            //console.log("c number:" + cNumber);

            var highlightRowColumnFrom;
            var highlightRowColumnUpto;
            if ($("#chkHeaderRow").prop("checked")) {
                highlightRowColumnFrom = 'R' + (+rFrom + startRange + 1) + "C" + cNumber; //This "+1" is added to start selection from selected row instead of previous row.
                highlightRowColumnUpto = 'R' + (+rFrom + startRange + totalRowsToSelect + 1) + "C" + cNumber; //This "+1" is added to select last row.
            }
            else {
                highlightRowColumnFrom = 'R' + (startRange + 1) + "C" + cNumber; // +1 as in without header row, need to start from 1 not from 0, Googlesheets select rows from 1 index not from 0, like R1C1
                highlightRowColumnUpto = 'R' + (startRange + totalRowsToSelect + 1) + "C" + cNumber; //This "+1" is added to select last row.
            }

            selectedChartA1NotationColumnsWithRows.push(highlightRowColumnFrom + ":" + highlightRowColumnUpto);
            dataRowsAnnotation.push(highlightRowColumnFrom + ":" + highlightRowColumnUpto);

        }//end of for loop.

        //console.log("SELECT TO: => " + JSON.stringify(oldSelectedCellsA1Notations));
        //console.log("SELECT TO HIGHLIGHT: => " + JSON.stringify(selectedChartA1NotationColumnsWithRows));
        if (selectedChartA1NotationColumnsWithRows.length > 0 && $('#dropdownSheets option:selected').text().trim() != "" && $('#dropdownSheets option:selected').text() != "Select Sheet") {

            if (oldSelectedCellsA1Notations.length > 0) {
                highlightedSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumnsWithRows, oldSelectedCellsA1Notations);
            }
            else {
                highlightedSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumnsWithRows, []);
            }

            oldSelectedCellsA1Notations = selectedChartA1NotationColumnsWithRows;
        }
        //else if (selectedChartA1NotationColumnsWithRows.length == 0 && $('#dropdownSheets option:selected').text().trim() != "" && $('#dropdownSheets option:selected').text() != "Select Sheet") {
        //    if (dataRowsAnnotation.length > 0)
        //    {
        //        highlightedSheetCells($('#dropdownSheets').val(), dataRowsAnnotation, []);
        //    }
        //}
        //highlightSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumns);
    }
    /*
    else {

        if ($('#dropdownSheets option:selected').text().trim() != "" && $('#dropdownSheets option:selected').text() != "Select Sheet") {
            highlightedSheetCells($('#dropdownSheets').val(), selectedChartA1NotationColumns, []);
        }
        //console.log(JSON.stringify(selectedChartA1NotationColumnsWithRows));
    }*/
}

function updateA1NotationColumn(startValue, newA1Notation) {
    var isNewEntry = true;
    //for (var counter = 0; counter < selectedChartA1NotationColumns.length; counter++)
    //{
    //    if (selectedChartA1NotationColumns[counter].startsWith(startValue))
    //    {
    //        selectedChartA1NotationColumns[counter] = newA1Notation;
    //        isNewEntry = false;
    //    } //end of if statement.
    //}//end of for loop.

    if (isNewEntry) {
        selectedChartA1NotationColumns.push(newA1Notation);
    }
}
function highlightedSheetCells(sheetName, selectedRange, oldSelectedRangeList) {
    if (!isSampleDataClicked || isEditModeClicked) {

        //console.log("highlightedSheetCells() => " + JSON.stringify(selectedRange));


        //if (!inSyncSheetWithAddonMode) {
        // console.log("highlightedSheetCells called + (inSyncSheetWithAddonMode=)" + inSyncSheetWithAddonMode);
        var isconnected = checkInternetConnection();
        if (!isconnected) {
            $(".se-pre-con").fadeOut("slow");
            showReconnectingOverlay();
            checkInternetHandler = setInterval(checkInternetConnection, 3000);
            return;
        }
        google.script.run.withSuccessHandler(function () {
            //$(".se-pre-con").fadeOut("slow");
        })
            .withFailureHandler(
            function (msg, element) {
                //$(".se-pre-con").fadeOut("slow");
                handleError(msg);
            }
            ).highlightSheetCells(sheetName, JSON.stringify(selectedRange), JSON.stringify(oldSelectedRangeList));
        //}
    }
}

function changeActiveSheet(sheetName) {
    google.script.run.selectActiveSheet(sheetName);




}

function changeActiveSheetForSync(sheetName) {
    //google.script.run.selectActiveSheet(sheetName);

    //console.log("changeActiveSheetForSync(sheetName)  called");
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    google.script.run.withSuccessHandler(function () {
        loadSelectedSheetA1Notation(sheetName);
        logUserActionIntoDatabase(selectedChartNameFromSelectChartUI + "-SelectSheetsDDL_Changed", "Charts");
        enableControl();

        syncChart();
        syncSheetDataWithAddonTimerHandler = setInterval(syncSheetDataWithGooglesheetAddonHandler, syncSheetDataWithAddonTimeSpan);
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).selectActiveSheet(sheetName);

}

function setDefaultsForDataSourceScreen() {
    if (isSampleDataClicked) {
        $('#addDimensionClick').css("display", "none");
        $('#addMeasureClick').css("display", "none");
        //$('.dimensionRemoveClass').css("display", "none");

        $("#selectDataSourceTitle").html("");//Explore Sample Data
        $("#selectDataSourceTitle").addClass("clickableExploreSampleData");
        $("#selectDataSourceTitle").css("margin", "0px 0px");
        $("#clickableExploreSampleData").css("display", "block");
    }
    else {
        $('#addDimensionClick').css("display", "block");
        $('#addMeasureClick').css("display", "block");
        //$('.dimensionRemoveClass').css("display", "block");
        var html = '<div style="float:left">Please select a sheet for&nbsp</div><div class="selectDataSourceTitle_ChartName" title="' + selectedchartDisplayName + '">' + selectedchartDisplayName + '</div>';
        $("#selectDataSourceTitle").html(html);
        $("#selectDataSourceTitle").removeClass("clickableExploreSampleData");
        $("#selectDataSourceTitle").css("margin", "7px 0px");
        $("#clickableExploreSampleData").css("display", "none");
    }
}

function CleanMeasureData(data) {

    var measureValueRegularExpression = /^[0-9]*[.]?[0-9]+$/;
    var currencyRegularExpression = /[$£€%Kr,]/g;
    for (var counter = 0; counter < data.length; counter++) {
        var nextData = data[counter];
        var dataColumnsLength = nextData.length;

        var measureColumnData = nextData[dataColumnsLength - 1];

        if (measureColumnData !== null) {
            //remove leading and trailing spaces.
            measureColumnData = measureColumnData.trim();
            measureColumnData = measureColumnData.replace(currencyRegularExpression, "");//removing currency symbols from measure column.
            //If after removing currency space exist
            measureColumnData = measureColumnData.trim();
        }
        if (measureColumnData === null || measureColumnData == undefined || measureColumnData == "") {
            measureColumnData = "0";
        }
        else if (measureColumnData.toLowerCase() == "true") {//If measure contains boolean value.
            measureColumnData = "1";
        }
        else if (measureColumnData.toLowerCase() == "false") {//If measure contains boolean value.
            measureColumnData = "0";
        }
        else if (!measureValueRegularExpression.test(measureColumnData)) // If measure column contains other than numeric value change it to 0.
        {
            measureColumnData = "0";
        }

        nextData[dataColumnsLength - 1] = measureColumnData;
    }
}

function formatMeasureCellValue(cellValue) {

    var measureValueRegularExpression = /^[0-9]*[.]?[0-9]+$/;
    var currencyRegularExpression = /[$£€%,]/g;

    if (cellValue == null) {
        cellValue = "";
    }

    var measureColumnData = cellValue.toString();
    //console.log("measureColumnData = " + measureColumnData);

    //if (measureColumnData != null) {
    //    //remove leading and trailing spaces.
    //    measureColumnData = measureColumnData;
    //    measureColumnData = measureColumnData.replace(currencyRegularExpression, "");//removing currency symbols from measure column.
    //    //If after removing currency space exist
    //    measureColumnData = measureColumnData;
    //}
    if (measureColumnData == null || measureColumnData == undefined || measureColumnData == "") {
        //console.log("1 condition true [" + measureColumnData);
        measureColumnData = "0";
    }
    else if (measureColumnData.toLowerCase() == "true") {//If measure contains boolean value.
        // console.log("2 condition true [" + measureColumnData);
        measureColumnData = "1";
    }
    else if (measureColumnData.toLowerCase() == "false") {//If measure contains boolean value.
        // console.log("3 condition true [" + measureColumnData);
        measureColumnData = "0";
    }
    //else if (!measureValueRegularExpression.test(measureColumnData)) // If measure column contains other than numeric value change it to 0.
    //{
    //    // console.log("3 condition true [" + measureColumnData);
    //    measureColumnData = "0";
    //}
    else {
        measureColumnData = measureCleansing(measureColumnData);
    }

    return +measureColumnData;
}
function measureCleansing(measureVal) {
    var cleanMeasure = measureVal.replace(/ /g, '').replace(/[.]+/g, '.').replace(/[-]+/g, '-').split("").filter(function (d) {
        if (d == "." || d == "-") {
            return true;
        }
        return !isNaN(d);
    }).join("");

    if (cleanMeasure.split(".").length > 2) {
        cleanMeasure = "." + cleanMeasure.replace(/[.]+/g, '');
    }
    if (cleanMeasure.split("-").length > 2) {
        cleanMeasure = "-" + cleanMeasure.replace(/[-]+/g, '');
    }

    return cleanMeasure;
}
function getTicksNumbers(min, max) {
    var ticksNumberArray = [];
    var distanceNumber = (max - min) / 7;
    if (distanceNumber > 0) {
        distanceNumber = Math.ceil(distanceNumber);
    }
    else {
        distanceNumber = 1;
    }
    for (var i = 0; i < 7; i++) {
        if (i == 0) {
            ticksNumberArray.push(min);
        }
        else if ((distanceNumber * i) + min > max - distanceNumber) {
            continue;
        }
        else {
            ticksNumberArray.push((distanceNumber * i) + min);

        }
    }
    return ticksNumberArray;
}

function generateSliderTicks(min, max) {
    var width = 290, height = 25;
    var margin = { top: 2, right: 15, bottom: 20, left: 6 };
    var widthAfterMargin = width - margin.left - margin.right;
    var heightAfterMargin = height - margin.bottom - margin.top;
    var tickNumberArray = getTicksNumbers(min, max);
    if (tickNumberArray[tickNumberArray.length - 1] != max) {
        tickNumberArray.push(max);
    }
    d3.select("#sliderTicks").html("");
    d3.select("#sliderTicks").append('svg')
        .attr('width', width)
        .attr('height', height);
    if (min == max) {
        max = "";
    }
    if (min > max) {
        min = min - 1;
        max = max + 1;
    }
    var data = [min, max];

    var svg = d3.select("#sliderTicks").select('svg').append('g')
        .attr('transform', 'translate(' + margin.left + ',' + margin.top + ')');

    var y = d3.scale.linear().range([0, widthAfterMargin]).domain(data);

    var xAxis = d3.svg.axis().scale(y).tickSize(0).tickValues(tickNumberArray).orient('bottom');//.ticks(5);

    svg.append('g')
        .attr('class', 'x axis')
        .attr('transform', 'translate(3, ' + margin.top + ')')
        .call(xAxis);

    // Remove ticks with less than 1
    svg.selectAll("text").style("opacity", function (d, i) {
        if (+d % 1 != 0) {
            return 0;
        }
        else {
            var format = d3.format(".2s");
            if (+d > 999) {
                d3.select(this).text(format(+d));
            }
            else {
                d3.select(this).text(+d.toFixed(0));
            }
            return 1;
        }
    });
    //Shift last tick from right to 8px left due cutting tick text problem 
    svg.select(".x.axis").selectAll('g')
        .each(function (d, i) {
            if (d == max) {
                var transform_g = $(this).attr("transform").replace("translate(", "").replace(")", "");
                var commaSeparated_g = transform_g.split(",");
                commaSeparated_g[0] = commaSeparated_g[0] - 7;
                $(this).attr('transform', 'translate(' + commaSeparated_g[0] + ',' + commaSeparated_g[1] + ')');
            }
        });
}

function processEmptyRowCells(dataWithoutFormating, dimensions, measures) {
    //1. remove row in which all columns are empty
    //2. handle measures with symbol
    // returned processed data
    //console.log("Before formating" + JSON.stringify(dataWithoutFormating));

    var newData = [];
    //var nullValueCount = 0;
    var columnNonEmptyValues = {};
    for (var i = 0; i < dataWithoutFormating.length; i++) {
        var row = dataWithoutFormating[i];
        // console.log("row " + JSON.stringify(row));
        var allEmpty = true;
        for (var j = 0; j < row.length; j++) {
            if (row[j] != '') {
                allEmpty = false;
                break;
            }
        }
        if (!allEmpty) {
            // handle empty dimension cell
            if (selectedChart == "SentimentTrendChart") {
                for (var k = 0; k < 1; k++) {

                    if (row[k] == '') {
                        row[k] = "NULL" + k;
                        //nullValueCount++
                    }

                    //console.log("dim = " + row[k]);
                }
            }
            else {

                for (var k = 0; k < dimensions.length; k++) {

                    if (row[k] == '') {
                        if (i == 0) {
                            row[k] = "NULL" + k;
                            columnNonEmptyValues[k] = row[k];
                        }
                        else {
                            row[k] = columnNonEmptyValues[k];//"NULL" + k;
                        }
                        //nullValueCount++
                    }
                    else {
                        columnNonEmptyValues[k] = row[k];
                    }
                    //console.log("dim = " + row[k]);
                }
            }
            // handle empty measure cell
            //console.log(dimensions.length + "  " + measures.length);
            var start = dimensions.length;
            for (var k = start; k < (dimensions.length + measures.length); k++) {
                //console.log("k " + k + " data " + row[k] + " length " + (dimensions.length + measures.length));
                if (row[k] == '') {
                    row[k] = 0;
                }
                else {
                    row[k] = +clearMeasureForAllCurrencies(row[k]);//formatMeasureCellValue(row[k]);
                }
                //console.log(row[k]);
            }
            newData.push(row);
        }
    }
    //console.log("After formating" + JSON.stringify(newData));
    return newData;
}

function hideAllElementsOnLoadMyChartsScreen() {
    $('#charticonsmenubar .separatorIcon').hide();
    $('#charticonsmenubar #iconMenuBreadcrumb').hide();
    $('#charticonsmenubar #divChartExpoChartsSearchBox').hide();
    $('#charticonsmenubar #divMyChartsSearchBox').hide();
    $('#charticonsmenubar .selectedChartNameDiv').hide();
    $('#ChartouterDiv').hide();
}
function openMyChartsContainerView(openingSource) {
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    hideAllElementsOnLoadMyChartsScreen();
    if (openingSource != undefined && openingSource != "" && openingSource.length > 0) {
        logUserActionIntoDatabase(openingSource, "Addon");
    }
    boldSelectedTopMenuOption("Bold", "normal", "normal", "normal", "normal", "normal");
    //$(".se-pre-con").fadeIn("slow");
    google.script.run.withSuccessHandler(function (savedChartsList) {
        populateMyChartsDiv(savedChartsList);
        //$(".se-pre-con").fadeOut("slow");
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).getMySavedChartsList();
    //showMyChartsContainerDiv();
}


function getCreationDateInMiliseconds(myChartCreationDateString) {
    //alert(myChartCreationDateString);
    // input date is in format /Date(1588942650617)
    var dateInMiliseconds = (myChartCreationDateString.split("/Date(")[1]).split(")/")[0];
    return dateInMiliseconds;
}

function giveMyChartsOrder(currentorder) {
    if (currentorder == "ascending") {
        var selectedSheetId = myChartsSheetId;
        $("#myChartsThumbnailsDiv").html("");
        for (var i = 0; i < mychartsArray.length; i++) {
            var isViewable = isViewableMyChartList(mychartsArray[i]);
            var currentMyChartsListChartCategory = '';
            var chartMetaJSON = JSON.parse(mychartsArray[i].ChartMetaJSON);
            if (chartMetaJSON != null) {
                currentMyChartsListChartCategory = chartMetaJSON.chartCategory;
            }
            if (mychartsArray[i].SheetId == selectedSheetId) {
                var thumbnailViewIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/view-chart-icon.png';
                var thumbnailRemoveIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/remove-chart-icon.png';
                var thumbnailInsertIcon = 'https://apps.polyvista.com/GooglesheetFeb2021/scripts/polyvista/new223/insert-icon.png';
                var thumbnailCreateChartIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/insert-chart-icon.png';
                var thumbnailAddIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/small-create-new-chart-icon.png';
                var thumbnaileditIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/edit-chart-icon.png';
                var thumbnailViewIcon_disabled = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/view-chart-disable-icon.png';
                var thumbnailInsertIcon_disabled = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/insert-chart-disable-icon.png';
                var thumbnailAddNewSampleSheet = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/add_samplesheet_data_chart.png';
                var selectedChartName = mychartsArray[i].ChartName;
                var selectedChartCustomName = JSON.parse(JSON.stringify(mychartsArray[i].ChartCustomName));
                //console.log(JSON.stringify(myCharts[i]));
                // As need latest state of chart that's why using updatedOn inplace of createdOn'
                //var dateInMiliseconds = getCreationDateInMiliseconds(JSON.stringify(myCharts[i].UpdatedOn) != "null" ? myCharts[i].UpdatedOn : myCharts[i].CreatedOn);

                var dateInMiliseconds = getCreationDateInMiliseconds(mychartsArray[i].CreatedOn);

                var selectedChartGUID = mychartsArray[i].ChartMetaGuid;
                var chartIconPath = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/Thumbnails/icon_' + mychartsArray[i].ChartName + '.png';

                var newThumbnailHTML = '<div class="RowOptionsDiv" id="thumbnailChartGuid_' + selectedChartGUID + '">'
                    + '<div class="ChartImageDiv">'
                    + '<img class="imagestyles" src="' + chartIconPath + '" alt="' + selectedChartName + '" title="' + selectedChartCustomName + '"/>'
                    + '</div>'
                    + '<div class="ChartOptionsLabels" >'
                    + '<span class="charttitle"  title="' + selectedChartCustomName + '">' + selectedChartCustomName + '</span>';

                newThumbnailHTML += '<div class="chartoptionsrowdiv CreateNewFromSavedChart" editableChartName="' + selectedChartCustomName + '" chartCategory="' + currentMyChartsListChartCategory + '" title="Create new chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnailCreateChartIcon + '" />'
                    + '<p class="createchart">Create New Chart</p>'
                    + '</div>';

                if (isViewable) {
                    newThumbnailHTML += '<div class="chartoptionsrowdiv openMySavedChart" editableChartName="' + selectedChartCustomName + '" title="View chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailViewIcon + '" />'
                        + '<p class="viewchart">View Chart</p>'
                        + '</div>';
                }
                else {
                    newThumbnailHTML += '<div style="cursor:default" class="chartoptionsrowdiv" editableChartName="' + selectedChartCustomName + '" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailViewIcon_disabled + '" />'
                        + '<p class="viewchart" style="color:#b8b8b8">View Chart</p>'
                        + '</div>';
                }

                newThumbnailHTML += '<div class="chartoptionsrowdiv editMySavedChart" editableChartName="' + selectedChartCustomName + '" title="Edit chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnaileditIcon + '" />'
                    + '<p class="editchart">Edit Chart</p>'
                    + '</div>'

                    + '<div class="chartoptionsrowdiv removeFromMySavedCharts" title="Remove chart from list" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" >'
                    + '<img class="imagestyles" src="' + thumbnailRemoveIcon + '" />'
                    + '<p class="removechart">Remove Chart</p>'
                    + '</div>';

                if (isViewable) {
                    newThumbnailHTML += '<div class="chartoptionsrowdiv insertChartIntoSheet" title="Insert chart into sheet" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailInsertIcon + '" />'
                        + '<p class="insertchart">Insert Chart in Sheet</p>'
                        + '</div>';
                }
                else {
                    newThumbnailHTML += '<div style="cursor:default" class="chartoptionsrowdiv" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailInsertIcon_disabled + '" />'
                        + '<p class="insertchart" style="color:#b8b8b8">Insert Chart in Sheet</p>'
                        + '</div>';
                }

                newThumbnailHTML += '<div class="chartoptionsrowdiv addSampleSheet" chartCategory="' + currentMyChartsListChartCategory + '" title="Click here to add a sheet with sample data and chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" style="margin-left:12px" src="' + thumbnailAddNewSampleSheet + '" />'
                    + '<p class="sampledata">Add Sample Chart Sheet</p>'
                    + '</div>';
                newThumbnailHTML += '</div>'
                    + ' </div>'
                    + '<div style="clear:both"></div>';

                // To Do List = Remove
                // newThumbnailHTML = newThumbnailHTML.replace("NaN", "");

                $("#myChartsThumbnailsDiv").append(newThumbnailHTML);
            }

        }
    }
    else {
        var selectedSheetId = myChartsSheetId;
        $("#myChartsThumbnailsDiv").html("");
        for (var i = mychartsArray.length - 1; i >= 0; i--) {
            var isViewable = isViewableMyChartList(mychartsArray[i]);
            var currentMyChartsListChartCategory = '';
            var chartMetaJSON = JSON.parse(mychartsArray[i].ChartMetaJSON);
            if (chartMetaJSON != null) {
                currentMyChartsListChartCategory = chartMetaJSON.chartCategory;
            }
            if (mychartsArray[i].SheetId == selectedSheetId) {
                var thumbnailViewIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/view-chart-icon.png';
                var thumbnailRemoveIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/remove-chart-icon.png';
                var thumbnailInsertIcon = 'https://apps.polyvista.com/GooglesheetFeb2021/scripts/polyvista/new223/insert-icon.png';
                var thumbnailCreateChartIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/insert-chart-icon.png';
                var thumbnailAddIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/small-create-new-chart-icon.png';
                var thumbnaileditIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/edit-chart-icon.png';
                var thumbnailViewIcon_disabled = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/view-chart-disable-icon.png';
                var thumbnailInsertIcon_disabled = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/insert-chart-disable-icon.png';
                var thumbnailAddNewSampleSheet = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/add_samplesheet_data_chart.png';
                var selectedChartName = mychartsArray[i].ChartName;
                var selectedChartCustomName = JSON.parse(JSON.stringify(mychartsArray[i].ChartCustomName));

                // As need latest state of chart that's why using updatedOn inplace of createdOn'
                //var dateInMiliseconds = getCreationDateInMiliseconds(JSON.stringify(myCharts[i].UpdatedOn) != "null" ? myCharts[i].UpdatedOn : myCharts[i].CreatedOn);
                var dateInMiliseconds = getCreationDateInMiliseconds(mychartsArray[i].CreatedOn);

                var selectedChartGUID = mychartsArray[i].ChartMetaGuid;
                var chartIconPath = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/Thumbnails/icon_' + mychartsArray[i].ChartName + '.png';

                var newThumbnailHTML = '<div class="RowOptionsDiv" id="thumbnailChartGuid_' + selectedChartGUID + '">'
                    + '<div class="ChartImageDiv">'
                    + '<img class="imagestyles" src="' + chartIconPath + '" alt="' + selectedChartName + '" title="' + selectedChartCustomName + '"/>'
                    + '</div>'
                    + '<div class="ChartOptionsLabels" >'
                    + '<span class="charttitle"  title="' + selectedChartCustomName + '">' + selectedChartCustomName + '</span>';

                newThumbnailHTML += '<div class="chartoptionsrowdiv CreateNewFromSavedChart" editableChartName="' + selectedChartCustomName + '" chartCategory="' + currentMyChartsListChartCategory + '" title="Create new chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnailCreateChartIcon + '" />'
                    + '<p class="createchart">Create New Chart</p>'
                    + '</div>';

                if (isViewable) {
                    newThumbnailHTML += '<div class="chartoptionsrowdiv openMySavedChart" editableChartName="' + selectedChartCustomName + '" title="View chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailViewIcon + '" />'
                        + '<p class="viewchart">View Chart</p>'
                        + '</div>';
                }
                else {
                    newThumbnailHTML += '<div style="cursor:default" class="chartoptionsrowdiv" editableChartName="' + selectedChartCustomName + '" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailViewIcon_disabled + '" />'
                        + '<p class="viewchart" style="color:#b8b8b8">View Chart</p>'
                        + '</div>';
                }

                newThumbnailHTML += '<div class="chartoptionsrowdiv editMySavedChart" editableChartName="' + selectedChartCustomName + '" title="Edit chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnaileditIcon + '" />'
                    + '<p class="editchart">Edit Chart</p>'
                    + '</div>'

                    + '<div class="chartoptionsrowdiv removeFromMySavedCharts" title="Remove chart from list" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" >'
                    + '<img class="imagestyles" src="' + thumbnailRemoveIcon + '" />'
                    + '<p class="removechart">Remove Chart</p>'
                    + '</div>';

                if (isViewable) {
                    newThumbnailHTML += '<div class="chartoptionsrowdiv insertChartIntoSheet" title="Insert chart into sheet" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailInsertIcon + '" />'
                        + '<p class="insertchart">Insert Chart in Sheet</p>'
                        + '</div>';
                }
                else {
                    newThumbnailHTML += '<div style="cursor:default" class="chartoptionsrowdiv" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                        + '<img class="imagestyles" src="' + thumbnailInsertIcon_disabled + '" />'
                        + '<p class="insertchart" style="color:#b8b8b8">Insert Chart in Sheet</p>'
                        + '</div>';
                }

                newThumbnailHTML += '<div class="chartoptionsrowdiv addSampleSheet" chartCategory="' + currentMyChartsListChartCategory + '" title="Click here to add a sheet with sample data and chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" style="margin-left:12px" src="' + thumbnailAddNewSampleSheet + '" />'
                    + '<p class="sampledata">Add Sample Chart Sheet</p>'
                    + '</div>';
                newThumbnailHTML += '</div>'
                    + ' </div>'
                    + '<div style="clear:both"></div>';

                // To Do List = Remove
                // newThumbnailHTML = newThumbnailHTML.replace("NaN", "");

                $("#myChartsThumbnailsDiv").append(newThumbnailHTML);
            }

        }
    }
}
//New Thing added
var windowHeight = $(this).height();
var windowWidth = $(this).width();

function populateMyChartsDiv(myChartsList) {
    if (myChartsList !== '[]') {
        $('#ChartouterDiv').show();
        $('#selectSheetDDLContainerDiv').show();
        $('#CreateChartContentDiv').show();
        $('#myChartsThumbnailsDiv').show();
        loadMyChartSheetsDDL(myChartsList);
        //$('#CreateChartContentDiv').css('padding-top', '10px');
        $('.borderLine').show();
        viewScreen("MyChartsScreen");
        boldSelectedTopMenuOption("Bold", "normal", "normal", "normal", "normal", "normal");
        isMySavedChartExists = true;
    }
    else {
        $(".se-pre-con").fadeOut("slow");
        $('#ChartouterDiv').show();
        $('#selectSheetDDLContainerDiv').hide();
        $('#myChartsThumbnailsDiv').hide();
        $('#CreateChartContentDiv').show();
        $('#CreateChartContentDiv').css('padding-top', '60px');
        $('.borderLine').hide();
        viewScreen("LandingScreen");
        boldSelectedTopMenuOption("normal", "normal", "normal", "normal", "Bold", "normal");
        isMySavedChartExists = false;
    }
}

function viewScreen(screenName) {
    selectedScreen = screenName;
    hideTooltip();
    clearTimeout(syncSheetDataWithAddonTimerHandler); // default stop synching timer
    $("#divDrawChart").removeClass('activeDrawButton');
    $("#startRowTextBox").prop("readonly", true);
    $("#endRowTextBox").prop("readonly", true);

    if (screenName == "LandingScreen") {
        $('#charticonsmenubar .separatorIcon').hide();
        $('#charticonsmenubar #iconMenuBreadcrumb').hide();
        $('#charticonsmenubar #divChartExpoChartsSearchBox').hide();
        $('#charticonsmenubar #divMyChartsSearchBox').hide();
        $('#charticonsmenubar .selectedChartNameDiv').hide();
        //$('.buttonsContainer').hide();
        $('.samplesheetdata').hide();
    }
    else if (screenName == "MyChartsScreen") {
        $('#charticonsmenubar .separatorIcon').hide();
        $('#charticonsmenubar #iconMenuBreadcrumb').hide();
        $('#charticonsmenubar #divChartExpoChartsSearchBox').hide();

        $('#charticonsmenubar #divMyChartsSearchBox').show();
        $('#charticonsmenubar .selectedChartNameDiv').hide();
        //$('.buttonsContainer').hide();
        $('.samplesheetdata').hide();
    }
    else if (screenName == "AllChartsScreen") {
        $('#charticonsmenubar .separatorIcon').show();
        $('#charticonsmenubar #iconMenuBreadcrumb').show();

        $('#charticonsmenubar .selectedChartNameDiv').hide();
        $('#charticonsmenubar #divChartExpoChartsSearchBox').show();
        $('#charticonsmenubar #divMyChartsSearchBox').hide();
        //$('.buttonsContainer').hide();
        $('.samplesheetdata').hide();
    }
    else if (screenName == "DataSourceScreen") {
        $('#charticonsmenubar .separatorIcon').show();
        $('#charticonsmenubar #iconMenuBreadcrumb').show();
        $('#charticonsmenubar #divChartExpoChartsSearchBox').hide();
        $('#charticonsmenubar #divMyChartsSearchBox').hide();
        $('#charticonsmenubar .selectedChartNameDiv').show();
        //$('.buttonsContainer').show();
        $('.samplesheetdata').show();
    }
    else if (screenName == "PricingScreen") {
        $('#charticonsmenubar .separatorIcon').show();
        $('#charticonsmenubar #iconMenuBreadcrumb').show();
        $('#charticonsmenubar #divChartExpoChartsSearchBox').hide();

        $('#charticonsmenubar #divMyChartsSearchBox').hide();

        $('#charticonsmenubar .selectedChartNameDiv').hide();
        //$('.buttonsContainer').hide();
        $('.samplesheetdata').hide();
    }
    else if (screenName == "SubscriptionScreen") {
        $('#charticonsmenubar .separatorIcon').show();
        $('#charticonsmenubar #iconMenuBreadcrumb').show();
        $('#charticonsmenubar #divChartExpoChartsSearchBox').hide();

        $('#charticonsmenubar #divMyChartsSearchBox').hide();

        $('#charticonsmenubar .selectedChartNameDiv').hide();
        //$('.buttonsContainer').hide();
        $('.samplesheetdata').hide();
    }
    else if (screenName == "ManageTeamUsers") {
        $('#charticonsmenubar .selectedChartNameDiv').hide();
        //$('.buttonsContainer').hide();
        $('#charticonsmenubar #iconMenuBreadcrumb').show();
        $('#charticonsmenubar .separatorIcon').show();
        $('.samplesheetdata').hide();
    }
}

function boldSelectedTopMenuOption(myChart, viewSubscription, pricingScreen, manageDomain, createNewChart, sampleSheetData) {
    $('.mychart').css("font-weight", myChart);
    $('.viewsubscription').css("font-weight", viewSubscription);
    $('.managedomain').css("font-weight", manageDomain);
    $('.viewPricing').css("font-weight", pricingScreen);
    $('.createnewchart').css("font-weight", createNewChart);
    $('.samplesheetdata').css("font-weight", sampleSheetData);
}

var newchart_added_updated_timer = false;

function startTimerToCheckNewChartAddedUpdated() {
    if (!newchart_added_updated_timer) {
        newchart_added_updated_timer = 1;
        getUserPackageDetail();
    }
}

var newchart_added_updated_timer_handler, timeSpaneToCheckMyChart = 5000, timerToOpenPaypalDynamicPackageURL, timeSpanToOpenPaypalDynamicPackageURL = 1000;

function chartAddedUpdatedInMyCharts() {
    if (localStorageAccessible) {
        if (window.localStorage.getItem("chartAddedUpdatedIntoMyChartList") != null && window.localStorage.getItem("chartAddedUpdatedIntoMyChartList") === "1") {
            openMyChartsView();
            newchart_added_updated_timer = 0;
            clearTimeout(newchart_added_updated_timer_handler);
            //window.localStorage.setItem("chartAddedUpdatedIntoMyChartList", "0");
            saveDateInTempStorage("chartAddedUpdatedIntoMyChartList", "0");
        }
        else {
            newchart_added_updated_timer_handler = setTimeout(chartAddedUpdatedInMyCharts, timeSpaneToCheckMyChart);
        }
    }
    else {
        if (storableObjectInTempStorage["chartAddedUpdatedIntoMyChartList"] != null && storableObjectInTempStorage["chartAddedUpdatedIntoMyChartList"] === "1") {
            openMyChartsView();
            newchart_added_updated_timer = 0;
            clearTimeout(newchart_added_updated_timer_handler);
            //window.localStorage.setItem("chartAddedUpdatedIntoMyChartList", "0");
            saveDateInTempStorage("chartAddedUpdatedIntoMyChartList", "0");
        }
        else {
            newchart_added_updated_timer_handler = setTimeout(chartAddedUpdatedInMyCharts, timeSpaneToCheckMyChart);
        }
    }
}

function openMyChartsView() {
    if (trialExpired) {
        return;
        //showTrialTopBar(true, licenceMessage, -1, undefined, undefined, undefined);
        //ShowMainMenuViews('licenseKeyContainer');
        //viewScreen("SubscriptionScreen");
    }
    else {
        $(".se-pre-con").fadeIn("slow");
        boldSelectedTopMenuOption("Bold", "normal", "normal", "normal", "normal", "normal");
        hideOtherDivs();
        goToHomePage();
        openMyChartsContainerView("ViewMyCharts-LandingPage");
    }
}

function hideOtherDivs() {
    $("#divChartExpoChartsSearchBox").hide();
    $('#CreateChartContentDiv').show();
    $('#DropdownChartsDiv').hide();
    $('#DataSourceDiv').hide();
    //$('#DataSourceDivHeaderRow').hide();
    $("#divChartExpoCharts").hide();
    //$("#divSankeyCharts").hide();
    $("#licenseKeyContainer").hide();
    $(".selectedChartNameDiv").show();
    //$(".buttonsContainer").show();
}

var savedChartsList;
function loadMyChartSheetsDDL(myChartsList) {
    //$('#dropdownlistSheets').append($('<option value="Select Sheet">Select Sheet</option>'));
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    google.script.run.withSuccessHandler(function (sheetsName) {
        if (sheetsName != null && sheetsName != undefined && sheetsName.length > 0) {
            // console.log("loadMyChartsSheetsDDL => " + JSON.stringify(sheetsName));
            loadMyChartSheetsList(sheetsName, myChartsList);
            $(".se-pre-con").fadeOut("slow");

            if (myChartsSheetId != "" && myChartsSheetName != "") {
                $(".myChartSelectSheetLabel").text(myChartsSheetName);
                populateMyChartsView(myChartsList, myChartsSheetId);
            }
            else {
                $(".myChartSelectSheetLabel").text(sheetsName[0].name);
                populateMyChartsView(myChartsList, sheetsName[0].id);
                myChartsSheetId = sheetsName[0].id;
                myChartsSheetName = sheetsName[0].name;
            }
            savedChartsList = myChartsList;
        }
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).readAllSheetsNamesWithId();
}

function loadMyChartSheetsList(sheetsList, myChartsList) {
    $("#ulSheetsList").empty();

    if (sheetsList.length > 0) {
        for (var counter = 0; counter < sheetsList.length; counter++) {
            $("#ulSheetsList").append($('<li class="chartexpo_GoogleSheetAddon_tileMenu_li_selectsheet" sheetId="' + sheetsList[counter].id + '" originaltext="' + sheetsList[counter].name + '">' + sheetsList[counter].name + '</li>'));
        }//end of for loop.

        $(".chartexpo_GoogleSheetAddon_tileMenu_li_selectsheet").on('click', function () {
            var selectedValue = $(this).attr("sheetId");

            myChartsSheetName = $(this).text();
            myChartsSheetId = selectedValue;

            $(".myChartSelectSheetLabel").text($(this).text());
            hideSearchSheetMenu();
            $('#addSheetClickContainer1').find('.imagecontainer').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/dropdown-icon.png');
            populateMyChartsView(myChartsList, selectedValue);
            event.stopPropagation();
        });
    }
}

function hideSearchSheetMenu() {
    $("#myDropdownSelectSheet").css("display", "none");
    //$("#myDropdownSelectSheet").removeClass("show");
}

function onMyChartsSheetDDLChange(sheetName) {
    var selectedSheetName = $(this).attr('id') == undefined ? sheetName : $(this).val();

    if (selectedSheetName != "Select Sheet") {
        $(".se-pre-con").fadeIn("slow");
        var selectedDDLSheetId = $('#dropdownlistSheets option:selected').val();
        populateMyChartsView(savedChartsList, selectedDDLSheetId);
        $(".se-pre-con").fadeOut("slow");
    }

}

function populateMyChartsView(myChartsList, selectedDDLSheetId) {
    var thumbnailViewIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/view-chart-icon.png';
    var thumbnailViewIcon_disabled = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/view-chart-disable-icon.png';
    var thumbnailRemoveIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/remove-chart-icon.png';
    var thumbnailInsertIcon = 'https://apps.polyvista.com/GooglesheetFeb2021/scripts/polyvista/new223/insert-icon.png';
    var thumbnailCreateChartIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/insert-chart-icon.png';
    var thumbnailInsertIcon_disabled = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/insert-chart-disable-icon.png';
    var thumbnailAddIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/small-create-new-chart-icon.png';
    var thumbnailAddNewSampleSheet = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/add_samplesheet_data_chart.png';
    var thumbnaileditIcon = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v3/edit-chart-icon.png';

    var myCharts = JSON.parse(myChartsList);
    myCharts = myCharts.filter(function (d) {
        if (tabularDataCharts_DimensionColumnName.hasOwnProperty(d.ChartName) == false) {
            return true;
        }
        else {
            var UpdatedOn_dateInMiliseconds = +getCreationDateInMiliseconds(d.UpdatedOn);
            if (UpdatedOn_dateInMiliseconds > googleSheetTabularDataLaunchDate) {
                return true;
            }
            return false;
        }
    });
    var selectedSheetId = "";
    mychartsArray = myCharts;
    currentOrder = "ascending";
    var chartRuleObject;
    // Empty recent list
    $("#myChartsThumbnailsDiv").html("");
    $('#orderMyCharts').attr("title", "Order by descending");
    $('#orderMyCharts').attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v41/sort-desc.png");
    for (var i = myCharts.length - 1; i >= 0; i--) {
        if (myCharts[i].SheetId == selectedDDLSheetId) {
            selectedSheetId = selectedDDLSheetId;
            $('#CreateChartContentDiv').css('padding-top', '10px');
            var selectedChartName = myCharts[i].ChartName;
            var selectedChartCustomName = JSON.parse(JSON.stringify(myCharts[i].ChartCustomName));
            var isViewable = isViewableMyChartList(myCharts[i]);
            var currentMyChartsListChartCategory = '';
            var chartMetaJSON = JSON.parse(mychartsArray[i].ChartMetaJSON);
            if (chartMetaJSON != null) {
                currentMyChartsListChartCategory = chartMetaJSON.chartCategory;
            }
            // As need latest state of chart that's why using updatedOn inplace of createdOn'
            //var dateInMiliseconds = getCreationDateInMiliseconds(JSON.stringify(myCharts[i].UpdatedOn) != "null" ? myCharts[i].UpdatedOn : myCharts[i].CreatedOn);

            var dateInMiliseconds = getCreationDateInMiliseconds(myCharts[i].CreatedOn);

            //console.log("Creation date = > " + JSON.parse(JSON.stringify(myCharts[i].CreatedOn)) + " miliseconds " + dateInMiliseconds); // new Date(1591961011957).toDateString()

            var selectedChartGUID = myCharts[i].ChartMetaGuid;
            var chartIconPath = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/Thumbnails/icon_' + myCharts[i].ChartName + '.png';

            var newThumbnailHTML = '<div class="RowOptionsDiv" id="thumbnailChartGuid_' + selectedChartGUID + '">'
                + '<div class="ChartImageDiv">'
                + '<img class="imagestyles" src="' + chartIconPath + '" alt="' + selectedChartName + '" title="' + selectedChartCustomName + '"/>'
                + '</div>'
                + '<div class="ChartOptionsLabels" >'
                + '<span class="charttitle"  title="' + selectedChartCustomName + '">' + selectedChartCustomName + '</span>';

            newThumbnailHTML += '<div class="chartoptionsrowdiv CreateNewFromSavedChart" editableChartName="' + selectedChartCustomName + '" chartCategory="' + currentMyChartsListChartCategory + '" title="Create new chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                + '<img class="imagestyles" src="' + thumbnailCreateChartIcon + '" />'
                + '<p class="createchart">Create New Chart</p>'
                + '</div>';

            if (isViewable) {
                newThumbnailHTML += '<div class="chartoptionsrowdiv openMySavedChart" editableChartName="' + selectedChartCustomName + '" title="View chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnailViewIcon + '" />'
                    + '<p class="viewchart">View Chart</p>'
                    + '</div>';
            }
            else {
                newThumbnailHTML += '<div style="cursor:default" class="chartoptionsrowdiv" editableChartName="' + selectedChartCustomName + '" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnailViewIcon_disabled + '" />'
                    + '<p class="viewchart" style="color:#b8b8b8">View Chart</p>'
                    + '</div>';
            }

            newThumbnailHTML += '<div class="chartoptionsrowdiv editMySavedChart" editableChartName="' + selectedChartCustomName + '" title="Edit chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" createdOn="' + dateInMiliseconds + '" >'
                + '<img class="imagestyles" src="' + thumbnaileditIcon + '" />'
                + '<p class="editchart">Edit Chart</p>'
                + '</div>'

                + '<div class="chartoptionsrowdiv removeFromMySavedCharts" title="Remove chart from list" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '" >'
                + '<img class="imagestyles" src="' + thumbnailRemoveIcon + '" />'
                + '<p class="removechart">Remove Chart</p>'
                + '</div>';

            if (isViewable) {
                newThumbnailHTML += '<div class="chartoptionsrowdiv insertChartIntoSheet" title="Insert chart into sheet" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnailInsertIcon + '" />'
                    + '<p class="insertchart">Insert Chart in Sheet</p>'
                    + '</div>';
            }
            else {
                newThumbnailHTML += '<div style="cursor:default" class="chartoptionsrowdiv" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                    + '<img class="imagestyles" src="' + thumbnailInsertIcon_disabled + '" />'
                    + '<p class="insertchart" style="color:#b8b8b8">Insert Chart in Sheet</p>'
                    + '</div>';
            }

            newThumbnailHTML += '<div class="chartoptionsrowdiv addSampleSheet" chartCategory="' + currentMyChartsListChartCategory + '" title="Click here to add a sheet with sample data and chart" chartType="' + selectedChartName + '"  chartGuid="' + selectedChartGUID + '"  createdOn="' + dateInMiliseconds + '" >'
                + '<img class="imagestyles" style="margin-left:12px" src="' + thumbnailAddNewSampleSheet + '" />'
                + '<p class="sampledata">Add Sample Chart Sheet</p>'
                + '</div>';
            newThumbnailHTML += '</div>'
                + ' </div>'
                + '<div style="clear:both"></div>';

            // To Do List = Remove
            // newThumbnailHTML = newThumbnailHTML.replace("NaN", "");

            $("#myChartsThumbnailsDiv").append(newThumbnailHTML);
        }
        else {
            if (selectedSheetId != selectedDDLSheetId) {
                $('#CreateChartContentDiv').css('padding-top', '60px');
            }
        }
    }
    loadSheets();
    //$('#CreateChartContentDiv').hide();
    //$('#ChartouterDiv').show();
}

function AddScript(src, id, _method, _execute) {
    $(".se-pre-con").fadeIn("slow");
    if (_execute == undefined) {
        _execute = false;
    }
    var alreadyExist = checkScriptAlreadyExist(src, id);
    if (!alreadyExist) {

        var script = document.createElement('script');
        script.setAttribute('src', src);
        script.setAttribute('type', 'text/javascript');
        script.setAttribute('id', id);
        script.onload = function () {
            if (_execute) {

                _method();
            }
        }
        document.getElementsByTagName('head')[0].appendChild(script);
    } else {
        if (_method != undefined) {
            _method();
        }
    }
}
function checkScriptAlreadyExist(src, id) {
    var element = document.getElementById(id);
    if (element != null) {
        var source = document.getElementById(id).src;
        if (source.indexOf(src) != -1)
            return true;
        else
            return false;
    }
    return false;
}

function setSelectRowRangeTextboxInitialValues(min, max) {
    if (!$("#chkHeaderRow").is(':checked')) {
        min = 1;
        $("#startRowTextBox").val(1);
        $("#endRowTextBox").val(2);
    }
    $("#startRowTextBox").attr("min", min + parseInt($("#txtBoxHeaderRow").val()));
    $("#startRowTextBox").attr("max", max + parseInt($("#txtBoxHeaderRow").val()));
    $("#endRowTextBox").attr("min", min + parseInt($("#txtBoxHeaderRow").val()));
    $("#endRowTextBox").attr("max", max + parseInt($("#txtBoxHeaderRow").val()));

    if (!($("#chkHeaderRow").is(':checked'))) {
        $("#startRowTextBox").val(1);
        $("#endRowTextBox").val(2);
    }
    else {
        $("#startRowTextBox").val(2);
        $("#endRowTextBox").val(3);
    }

    maxNumberOfRecordsInSelectedSheet = max;// maximum number of records set
}

function resetDatasourceScreen() {
    //resetting datasource screen.
    $('.metric').remove();
    //$('.dimension').not(':first').remove();
    $('.dimension').remove();
    $('#dropDownDimensions0').empty();
    $('#dropdownSheets').empty();
    $("#txtBoxHeaderRow").val(1);
    $("#txtBoxHeaderRow").removeAttr("disabled");
    $('.dropdownchangeclass').removeAttr("disabled");

    $('.slidecontainer').css('display', 'block');
    $('#selectedRowsDiv').css('display', 'block');
    $('#spanSelectRowRange').css('display', 'block');

    selectedChartColumns = [];
    selectedChartA1NotationColumns = [];
    sheetA1NotationDetails = [];
    datasourceColumnsWithIndex = [];
    dataSourceColumns[0] = [];
    dataSourceColumnsAll[0] = [];
    actualSheetData = [];
    sheetRecords = [];

    dataSourceDuplicateColumns = [];
    //updating slider value.
    updateSliderValue();
}

function setSelectRowRangeTextboxValues(startFrom, endTo) {
    //$("#startRowTextBox").attr("min",);
    //$("#startRowTextBox").attr("max", );
    //$("#startRowTextBox").attr("min",);
    //$("#startRowTextBox").attr("max", );

    if (parseInt($("#txtBoxHeaderRow").val()) == 0) {
        startFrom = parseInt($("#txtBoxHeaderRow").val()) + startFrom + 1;
        endTo = parseInt($("#txtBoxHeaderRow").val()) + endTo;
    }
    else {
        startFrom = parseInt($("#txtBoxHeaderRow").val()) + startFrom + 1;
        endTo = parseInt($("#txtBoxHeaderRow").val()) + (endTo == 0 ? 1 : endTo);
    }
    $("#startRowTextBox").val(startFrom);
    if (parseInt($("#txtBoxHeaderRow").val()) == 0) {
        $("#endRowTextBox").val(endTo);
    }
    else {
        $("#endRowTextBox").val(endTo);
    }
}

function autoHideMessageDialog() {
    setTimeout(function () {
        $(".error-overlay").fadeOut("slow");
    }, 8000);
    setTimeout(function () {
        $(".error_dialog").fadeOut("slow");
    }, 8000);
}

function showSavingTopStrip() {

    // $("#savingTopStrip").show();
    //$(".se-pre-con").fadeIn("slow");
}
function hideSavingTopStrip() {
    // $("#savingTopStrip").hide();
    // $(".se-pre-con").fadeOut("slow");
}
function hideTooltip() {
    $(".info").html("");
}

function convertToNumberingScheme(number) {
    var baseChar = ("A").charCodeAt(0),
        letters = "";

    do {
        number -= 1;
        letters = String.fromCharCode(baseChar + (number % 26)) + letters;
        number = (number / 26) >> 0; // quick `floor`
    } while (number > 0);
    return letters;
}

function getDataObjectWithFilteredColumns(dimensions, measures, sheetRecordsIndex) {

    var filtereData = [];
    for (var counter = 0; counter < dimensions.length; counter++) {
        var columnIndex = getColumnIndexFromSheetData(dimensions[counter]);
        if (columnIndex >= 0 && sheetRecords.length > sheetRecordsIndex && sheetRecords[sheetRecordsIndex].length > columnIndex) {
            filtereData.push(sheetRecords[sheetRecordsIndex][columnIndex]);
        }//end of if statement.
    } //end of for loop.
    for (var counter = 0; counter < measures.length; counter++) {
        var columnIndex = getColumnIndexFromSheetData(measures[counter]);
        if (columnIndex >= 0 && sheetRecords.length > sheetRecordsIndex && sheetRecords[sheetRecordsIndex].length > columnIndex) {
            filtereData.push(sheetRecords[sheetRecordsIndex][columnIndex]);
        }//end of if statement.
    } //end of for loop.
    return filtereData;
}

function getColumnIndexFromSheetData(dimensionName) {
    //SheetColumnIndex ,ColumnName
    for (var counter = 0; counter < datasourceColumnsWithIndex.length; counter++) {
        if (dimensionName == datasourceColumnsWithIndex[counter].ColumnName) {
            return datasourceColumnsWithIndex[counter].SheetColumnIndex;
        }//end of if statement.
    }//end of for loop.
    return -1;
}

function getDatasourceScreenLatestState() {

    // google.script.run.showChartViewerInDialog();
    // var chartRuleObject = getSelectedChartRulesObject();
    // selectedchartDisplayName = chartRuleObject.ChartDisplayName;

    var drawChartObject = { ChartName: '', Sheet: '', HeaderRow: 0, Dimensions: [], Measures: [], RowStartIndex: 0, RowLastIndex: 0 };
    var dimensions = [];
    var measures = [];
    drawChartObject.ChartName = selectedChart;
    drawChartObject.Sheet = $('#dropdownSheets').val();
    drawChartObject.HeaderRow = parseInt($('#txtBoxHeaderRow').val());
    $(".dimension > div:first-child").each(function (index) {
        var isColumnDeleted = $(this).attr('isdeletedcolumn');
        //if (isColumnDeleted != "true") {
        dimensions.push($(this).text());
        //}
    });
    drawChartObject.Dimensions = dimensions;
    $(".metric > div:first-child").each(function (index) {
        var isColumnDeleted = $(this).attr('isdeletedcolumn');
        //if (isColumnDeleted != "true") {
        measures.push($(this).text());
        // }
    });
    drawChartObject.Measures = measures;

    if (isSampleDataClicked) {
        dimensions = [];
        measures = [];
        $(".dimension > select").each(function (index) {
            dimensions.push($(this).val());
        });
        $(".metric > select").each(function (index) {
            measures.push($(this).val());
        });

        drawChartObject.Dimensions = dimensions;
        drawChartObject.Measures = measures;
    }

    drawChartObject.RowLastIndex = maxValue;

    tableHeaderRow = [[]];
    var tableBodyRows = [];
    tableHeaderRow[0] = dimensions.concat(measures);

    if (isSampleDataClicked) {
        var dataRows = getDateFromTempStorage("dataRows");
        if (dataRows != null && dataRows.length > 0) {
            tableBodyRows = JSON.parse(dataRows);
        }
    }
    else {
        var headerRowNumber = parseInt($("#txtBoxHeaderRow").val());

        for (var counter = (startSliderRange - 1); counter < (drawChartObject.RowLastIndex); counter++) {
            if (counter < 0) {
                counter = 0;
            }
            if (counter < drawChartObject.RowLastIndex) { //if counter value is less than slider value.
                var filteredDataObject = getDataObjectWithFilteredColumns(dimensions, measures, counter);

                tableBodyRows.push(filteredDataObject);
            }
        }

        tableBodyRows = processEmptyRowCells(tableBodyRows, dimensions, measures);

    }

    var rowFrom;
    var rowTo;
    if ($("#chkHeaderRow").prop("checked")) {
        rowFrom = +$("#startRowTextBox").val() - 2;
        rowTo = +$("#endRowTextBox").val() - 2;
    }
    else {
        rowFrom = +$("#startRowTextBox").val();
        rowTo = +$("#endRowTextBox").val();
    }

    var finalState = {
        useHeaderRow: "", fileName: "", fileId: "", sheetName: drawChartObject.Sheet, sheetId: "",
        headerRow: tableHeaderRow, dataRows: tableBodyRows, a1NotationInformation: oldSelectedCellsA1Notations,
        "headerRowNumber": $("#txtBoxHeaderRow").val(), "dataRowFrom": rowFrom, "dataRowTo": rowTo,
        "headerRowAnnotation": JSON.stringify(headerRowAnnotation), "dataRowsAnnotation": JSON.stringify(dataRowsAnnotation),
        "selectedDimensions": JSON.stringify(drawChartObject.Dimensions), "selectedMeasures": JSON.stringify(drawChartObject.Measures)
    };
    //alert("on save " + $("#txtBoxHeaderRow").val());
    //console.log("on save " + $("#txtBoxHeaderRow").val());

    return finalState;
}

function syncSheetDataWithGooglesheetAddonHandler() {
    var selectedSheet = $('#dropdownSheets').val();
    var selectedHeaderRow = $("#txtBoxHeaderRow").val();

    if (+selectedHeaderRow == 0) {
        selectedHeaderRow = -1;
    }

    if (selectedSheet !== "Select Sheet") {
        syncSheetDataWithGooglesheetAddon(selectedSheet, selectedHeaderRow);
    }
}

var inSyncSheetWithAddonMode = false;

function syncSheetDataWithGooglesheetAddon(sheetName, headerRow) {
    //Get data of selected sheet.
    //showSavingTopStrip();
    var isconnected = checkInternetConnection();
    if (!isconnected) {
        $(".se-pre-con").fadeOut("slow");
        showReconnectingOverlay();
        checkInternetHandler = setInterval(checkInternetConnection, 3000);
        return;
    }
    inSyncSheetWithAddonMode = true;
    google.script.run.withSuccessHandler(function (sheetDataWithHeaderRowDetail) {
        if (sheetDataWithHeaderRowDetail != null && sheetDataWithHeaderRowDetail != undefined && sheetDataWithHeaderRowDetail.length > 0) {
            //console.log("syncSheetDataWithGooglesheetAddon=> " + sheetDataWithHeaderRowDetail);
            var newDataWithHeaderRowDetail = JSON.parse(sheetDataWithHeaderRowDetail);

            // if sheet is removed then moved to user default data source screen
            if (newDataWithHeaderRowDetail.sheetData == "SELECTED_SHEET_REMOVED") {
                // its mean selected sheet removed. now moved to user data source screen again with default state
                // stop timer
                // set syncMode add
                // reset all variables
                syncMode = "Add";
                console.log("SELECTED_SHEET_REMOVED");
                clearTimeout(syncSheetDataWithAddonTimerHandler); // default stop synching timer
                initializeDataSourceScreenWithDefaultState();
                return;
            }

            var oldNumberOfRows = sheetRecords.length;
            // update global objects
            actualSheetData = newDataWithHeaderRowDetail.sheetData;
            //hideSavingTopStrip();

            var processedElement = [];
            var newHeaderColumns = newDataWithHeaderRowDetail.headerRowColumnsDetail.columnNames
            dataSourceDuplicateColumns = [];

            processNewHeaderColumns(newHeaderColumns, processedElement, "true");
            //console.log("latest header row" + newDataWithHeaderRowDetail.headerRowColumnsDetail.headerRowActualData);
            fillDatasourceColumnsAndSheetDataAfterSyncOperation(actualSheetData);
            updateDataSourceScreenAfterAutoSync(dataSourceColumnsAll[0], actualSheetData.length, oldNumberOfRows);

            // no need to sync in db on every sheet to add-on sync while do it only on explicit actions like add/remove dropdown etc.
            //syncSelectedChartAtServer(selectedchartDisplayName, syncMode, synchedChartGUID, $('#chkHeaderRow').prop("checked"), selectedChartCategory, "", "", getDatasourceScreenLatestState());
        }
        else {
            //$('#divDrawChart').css('cursor', 'pointer');
            //hideSavingTopStrip();
        }
        inSyncSheetWithAddonMode = false;
    })
        .withFailureHandler(
        function (msg, element) {
            //hideSavingTopStrip();
            inSyncSheetWithAddonMode = false;
            handleError(msg);
        }).syncSheetDataWithAddon(sheetName, headerRow);
}

function updateDataSourceScreenAfterAutoSync(newHeaderColumns, numberOfRows, oldNumberOfRows) {
    //var newHeaderColumns = [['0', '2', '4', '6', 'ght']];
    //var numberOfRows = 10;
    if (newHeaderColumns.length > 0) {
        $(".dimension > div:first-child").each(function (index) {
            var columnText = $(this).find('div').text();
            setHtmlAttributesAfterSyncOperation(this, newHeaderColumns, columnText);

        }); //end of each loop.

        $(".metric > div:first-child").each(function (index) {
            var columnText = $(this).find('div').text();
            setHtmlAttributesAfterSyncOperation(this, newHeaderColumns, columnText);
        }); //end of each loop.


        var headerRowNumber = parseInt($("#txtBoxHeaderRow").val());
        if ($("#chkHeaderRow").prop("checked")) {
            if ((numberOfRows - headerRowNumber) != oldNumberOfRows) {
                updateSliderValue($('#startRowTextBox').val(), $('#endRowTextBox').val());
            }
        }
        else {
            if (numberOfRows != oldNumberOfRows) {
                updateSliderValue($('#startRowTextBox').val(), $('#endRowTextBox').val());
            }
        }

        if (chartRequiredNoOfDimAndMetricsChosen()) {
            $("#divDrawChart").addClass('activeDrawButton');
            //$('#divDrawChart').css("color", "#F37A2D");
            //$('#divDrawChart').css("background-color", "white");
            $("#startRowTextBox").prop("readonly", false);
            $("#endRowTextBox").prop("readonly", false);
        }
        else {
            $("#divDrawChart").removeClass('activeDrawButton');
            $('#divDrawChart').css("color", "#B8B8B8");
            $("#startRowTextBox").prop("readonly", true);
            $("#endRowTextBox").prop("readonly", true);
        }
    }//end of if statement.
}

function setHtmlAttributesAfterSyncOperation(obj, newHeaderColumns, columnText) {
    var isRenderedColumnExistInNewHeaderList = checkRenderedColumnInNewHeaderList(newHeaderColumns, columnText);
    //If it exist in new list then no need to do any operation.
    if (!isRenderedColumnExistInNewHeaderList) {
        //If rendered column does not exist in newHeaderList then mark it deleted and set its background.
        $(obj).attr("isdeletedcolumn", "true");
        $(obj).css("background-color", "#f08080");
        $(obj.parentElement).removeAttr("draggable");

    } //end of if statement.
    else {
        //If rendered column was already marked deleted but now it is added in sheet.
        $(obj).attr("isdeletedcolumn", "false");
        $(obj).css("background-color", "white");
        $(obj.parentElement).attr("draggable", "true");
    }
}

function checkRenderedColumnInNewHeaderList(newHeaderColumns, columnText) {
    for (var counter = 0; counter < newHeaderColumns.length; counter++) {
        if (newHeaderColumns[counter].toLowerCase() == columnText.toLowerCase()) {
            return true;
        }
    }
    return false;
}

function updateSliderValue(defaultStartHandler, defaultEndHandler) {
    //setting range control max value.
    //alert("sheetRecords.length: " + sheetRecords.length + ", defaultStartHandler: " + defaultStartHandler + ", defaultEndHandler: " + defaultEndHandler);
    //alert(sheetRecords.length);
    var startRange = parseInt($("#txtBoxHeaderRow").val());

    // check applied as in case of header less row, we need to start ticks from 1 not from 0
    if ($("#chkHeaderRow").is(':checked') == false) {
        startRange = 1;
    }

    if (defaultEndHandler != undefined && defaultEndHandler != 0 && defaultStartHandler != undefined && defaultStartHandler != 0) {
        setMaximumSliderValue(sheetRecords.length, defaultStartHandler - startRange, defaultEndHandler - startRange);
    }
    else {
        setMaximumSliderValue(sheetRecords.length, defaultStartHandler, defaultEndHandler);
    }
    //$('#DataSourceMultiRangeSlider').html('');
    //generateSlider("DataSourceMultiRangeSlider", multiSliderMaxValue, sliderCallBack);
    //$("#maxRangeValue").html(sheetRecords.length);
    if (startRange < 1) {
        startRange = 0;
    }

    if ($("#chkHeaderRow").is(':checked') == false) {
        generateSliderTicks(startRange, sheetRecords.length + (startRange - 1));
    }
    else {
        generateSliderTicks(startRange + 1, sheetRecords.length + (startRange));
    }
    //setSelectRowRangeTextboxValues(0, sheetRecords.length);
    //setSelectRowRangeTextboxInitialValues(0, sheetRecords.length);
}

function getNonEmptyRowFromSheetData(sheetData) {
    for (var counter = 0; counter < sheetData.length; counter++) {
        var rowData = sheetData[counter];
        if (rowData != null) {
            rowData = rowData.join('');
            //console.log("row data: " + rowData);
            if (rowData != "") {
                return { rowNumber: (counter + 1), rowData: sheetData[counter] };
            }
        }
    }//end of for loop.
    return null;
}

function getSelectedChartDisplayName() {
    for (var counter = 0; counter < chartRulesObject.length; counter++) {
        if (chartRulesObject[counter].ChartName == selectedChart)
            return chartRulesObject[counter];
    }
}

function fillDatasourceColumnsAndSheetDataAfterSyncOperation(sheetData) {
    //sheetData = [["Topic", "Sub topic-1", "Sub topic-2", "Count"],["Daily Food Sales", "Breakfast", "Waffles", "80"], ["Daily Food Sales", "Breakfast", "Eggs", "60"], ["Daily Food Sales", "Breakfast", "Pancakes", "45"], ["Daily Food Sales", "Breakfast", "Tea", "30"], ["Daily Food Sales", "Lunch", "Salid", "100"], ["Daily Food Sales", "Lunch", "Sandwich", "80"], ["Daily Food Sales", "Lunch", "Soup", "50"], ["Daily Food Sales", "Lunch", "Pie", "35"]];
    var headerRow = $("#txtBoxHeaderRow").val();
    //dataSourceColumns[0] = [];
    //dataSourceColumnsAll[0] = [];
    // dataSourceDuplicateColumns = [];
    sheetRecords = [];
    headerRow = parseInt(headerRow) - 1;

    var headerColumns = [];
    var sheetA1Notation = '';
    //headerRow will be - 1 when $("#txtBoxHeaderRow").val() = 0
    for (var counter = headerRow; counter < sheetData.length; counter++) {

        if (counter == -1)//If there is no header column.
        {
            //console.log("sheet data" + JSON.stringify(sheetData[0]));
            var rowObject = getNonEmptyRowFromSheetData(sheetData);

            var sheetA1NotationWithNoHeader = [];
            if (rowObject != null && rowObject.rowData != null) {
                var nonEmptyRow = rowObject.rowData;
                for (var index = 0; index < nonEmptyRow.length; index++) {
                    if (nonEmptyRow[index] != null && nonEmptyRow[index] != "") {
                        headerColumns.push("Column " + convertToNumberingScheme((parseInt(index) + 1)));
                    }
                    else {
                        headerColumns.push("");
                    }
                    sheetA1NotationWithNoHeader.push("R" + rowObject.rowNumber + "C" + (parseInt(index) + 1));
                }//end of for loop.
                sheetA1Notation = sheetA1NotationWithNoHeader;
                //dataSourceColumns[0] = removeEmptyColumns(headerColumns, sheetA1NotationWithNoHeader); //Adding header row.    
                //dataSourceColumnsAll[0] = headerColumns;
            }
        }
        else {
            var rowData = sheetData[counter];
            if (rowData != null) {
                rowData = rowData.join('');
                if (counter == parseInt(headerRow)) {
                    sheetA1Notation = '';
                    if (sheetA1NotationDetails.length > counter) {
                        sheetA1Notation = sheetA1NotationDetails[counter];
                    }
                    //dataSourceColumns[0] = removeEmptyColumns(sheetData[counter], sheetA1Notation); //Adding header row.
                    //dataSourceColumnsAll[0] = sheetData[counter];
                }
                else {
                    sheetRecords.push(sheetData[counter]);
                }
            }
        }
    }//end of for loop.    
}

function removeEmptyColumns(dsColumns, sheetA1Notation) {
    var filterDatasourceColumns = [];
    datasourceColumnsWithIndex = [];
    for (var counter = 0; counter < dsColumns.length; counter++) {
        if (dsColumns[counter] != null && dsColumns[counter] != "") {
            filterDatasourceColumns.push(dsColumns[counter]);
            datasourceColumnsWithIndex.push({
                SheetColumnIndex: counter, ColumnName: dsColumns[counter],
                A1NotationDetail: (sheetA1Notation.length > counter ? sheetA1Notation[counter] : ""),
                A1NotationInRCFormat: (sheetA1Notation.length > counter ? sheetA1Notation[counter] : "")
            });
        }//end of if statement.
    } //end of for loop.
    return filterDatasourceColumns;
}
function chartRequiredNoOfDimAndMetricsChosen() {
    var chartRuleObject = getSelectedChartRulesObject();
    var dimensions = [];
    var measures = [];
    if (chartRuleObject != null) {
        //If any of the column is in deleted state then disable createchart button.
        var isDeletedColumnExist = false;
        $(".dimension > div:first-child").each(function (index) {
            var isColumnDeleted = $(this).attr('isdeletedcolumn');
            if (isColumnDeleted == "true") {
                isDeletedColumnExist = true;
            }
            dimensions.push($(this).text());
        });
        if (isDeletedColumnExist) {
            return false;
        }

        $(".metric > div:first-child").each(function (index) {
            var isColumnDeleted = $(this).attr('isdeletedcolumn');
            if (isColumnDeleted == "true") {
                isDeletedColumnExist = true;
            }
            measures.push($(this).text());
        });
        if (isDeletedColumnExist) {
            return false;
        }
        var selectedDataRows = +($("#rangeValue").html());

        if (dimensions.length < chartRuleObject.MinDim || measures.length < chartRuleObject.MinMetric || selectedDataRows < 1) {
            // required dimensions & measures not chosen
            // add check of minimum number of rows as well
            return false;
        }
    }
    return true;
}


function getSelectedChartRulesObject() {
    for (var counter = 0; counter < chartRulesObject.length; counter++) {
        if (selectedChartNameFromSelectChartUI != "" && chartRulesObject[counter].ChartName == selectedChartNameFromSelectChartUI)
            return chartRulesObject[counter];
    }
    return null;
}

function onMyChartEditButtonClicked(editMyChartObject) {
    //alert("onMyChartEditButtonClicked " + editMyChartObject.sheetName);
    //console.trace();
    isEditModeClicked = true;
    isEditModeLoadingFromMyChartClick = true;
    isSampleDataClicked = false;
    $('#btnDrawChartFromSampleData').removeClass('tabButtonActive');
    $('#btnDrawChartFromSheetData').removeClass('tabButtonActive');

    $("#divDrawChart").addClass('activeDrawButton');
    $("#startRowTextBox").prop("readonly", false);
    $("#endRowTextBox").prop("readonly", false);
    $('#divDrawChart').css("color", "#F37A2D");
    $('#divDrawChart').css("background-color", "white");
    $('#btnDrawChartFromSheetData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sheet-data-active.png');
    $('#btnDrawChartFromSheetData').find('.tabContainer').css('color', '#F37A2D');

    $('#btnDrawChartFromSampleData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sample-data-org.png');
    $('#btnDrawChartFromSampleData').find('.tabContainer').css('color', 'black');

    if (editMyChartObject != null) {
        selectedChartNameFromSelectChartUI = editMyChartObject.selectedChart;
        selectedChartCategory = editMyChartObject.selectedChartCategory;
        selectedChart = selectedChartNameFromSelectChartUI;
        loadDataSourceContainer(editMyChartObject);

    }
}

function loadDataSourceContainer(editMyChartObject) {
    //alert("isEditModeClicked = " + isEditModeClicked+"  " + JSON.stringify(editMyChartObject));
    //$('#DataSourceDiv').height(windowHeight - Heightcharticonsmenubar);
    //$('#DataSourceDivHeaderRow').css('display', 'block');
    $('#DataSourceDiv').css('display', 'block');
    $('#divChartExpoCharts').css('display', 'none');
    $('#divChartExpoChartsSearchBox').css('display', 'none');
    //$("#divSankeyCharts").css("display", "none");
    $("#ChartouterDiv").css("display", "none");

    // set name and icon of selected chart
    // selectedChartNameFromSelectChartUI


    var chartIconPath = 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/Thumbnails/icon_' + selectedChartNameFromSelectChartUI + '.png';

    $('.selectedChartIcon').css("cursor", "pointer");
    $('.selectedChartIcon').attr("src", chartIconPath);
    $('.selectedChartIcon').attr("title", "View charts list");
    $('.selectedChartIcon').css({ "width": "18px", "height": "18px" });

    //setting adddimension and addmeasure buttons text.
    var chartRuleObject = getSelectedChartRulesObject();
    selectedchartDisplayName = chartRuleObject.ChartDisplayName;
    $('.selectedChartTitle').html(getChartNameWithEllipses(selectedchartDisplayName));
    $('.selectedChartTitle').attr("title", selectedchartDisplayName);

    if (chartRuleObject != null) {
        $('#addDimensionClick').find('span').html("Add new " + chartRuleObject.DimensionText);
        $('#addMeasureClick').find('span').html("Add new " + chartRuleObject.MeasureText);

        if (chartRuleObject.MaxDim == 1) {
            $('#spanDimension').html("Please select a column for " + chartRuleObject.DimensionsBoxTitle + ":");
        }
        else {
            $('#spanDimension').html("Please select columns for " + chartRuleObject.DimensionsBoxTitle + ":");
        }
        if (chartRuleObject.MaxMetric == 1) {
            $('#spanMeasure').html("Please select a column for " + chartRuleObject.MeasuresBoxTitle + ":");
        }
        else {
            $('#spanMeasure').html("Please select columns for " + chartRuleObject.MeasuresBoxTitle + ":");
        }
    }

    $('#addDimensionClick').css("display", "block");
    $('#addMeasureClick').css("display", "block");

    if (isSampleDataClicked) {
        fillDatasourceScreenWithSampleDataDetails();
        // 1. hide add measure and dimension buttons
        // 2. change data source title and link it with data viewer
        // 3. 
        setDefaultsForDataSourceScreen();
    }
    else if (isEditModeClicked) {
        setDefaultsForDataSourceScreen();
        fillDatasourceScreenInEditMode(editMyChartObject); // This method populate datasource screen UI
    }
    else {
        resetDatasourceScreen();
        //attaching headerrow event.
        var txtHeaderRow = document.getElementById('txtBoxHeaderRow');
        txtHeaderRow.addEventListener("change", onHeaderChange);
        $(".se-pre-con").fadeIn("slow");
        //loading sheets.
        loadSheets();
        attachClickEvent('dimensionRemoveClass', 'dimension');
    }
}

function fillDatasourceScreenWithSampleDataDetails() {
    isSampleDataClicked = true;
    isEditModeClicked = false;
    isEditModeLoadingFromMyChartClick = false;
    var processedObject = getProcessedDataObjectForChartViewer("sampledata");
    storeDataInLocalStorage(processedObject);
    loadDatasourceScreenWithSampleData(processedObject);
}

function fillDatasourceScreenInEditMode(editMyChartObject) {
    //alert(editMyChartObject.headerRow);
    // Give new name to newly added sheet in case of edit mode
    newlyAddedSheetNameOnEditMode = "MyChart_" + createGuid();
    $('#btnDrawChartFromSheetData').addClass('tabButtonActive');
    $('#btnDrawChartFromSampleData').removeClass('tabButtonActive');

    $('#btnDrawChartFromSheetData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sheet-data-active.png');
    $('#btnDrawChartFromSheetData').find('.tabContainer').css('color', '#F37A2D');

    $('#btnDrawChartFromSampleData').find('.tabImg').attr('src', 'https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/sample-data-org.png');
    $('#btnDrawChartFromSampleData').find('.tabContainer').css('color', 'black');

    storeDataInLocalStorage(editMyChartObject);
    loadDatasourceScreenInEditMode(editMyChartObject);
}

function loadDatasourceScreenInEditMode(editMyChartObject) {
    resetDatasourceScreen();
    //attaching headerrow event.
    $("#txtBoxHeaderRow").removeAttr("disabled");
    $("#txtBoxHeaderRow").val(editMyChartObject.headerRowNumber);

    if (editMyChartObject.useHeaderRow.toString() == "true") {
        $("#chkHeaderRow").prop("checked", "checked");
        $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-active.png");
    }
    else {
        //console.log('fff:' + editMyChartObject.useHeaderRow);
        $("#chkHeaderRow").removeAttr("checked");
        //$('#chkHeaderRow').attr("disabled", "true");
        $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-deactive.png");
    }

    //$('#chkHeaderRow').attr("disabled", "true");

    var txtHeaderRow = document.getElementById('txtBoxHeaderRow');
    txtHeaderRow.addEventListener("change", onHeaderChange);
    //loading sheets in edit mode.
    $(".se-pre-con").fadeIn("slow");
    //loading sheets in edit mode.        
    loadSheetsInEditMode(editMyChartObject);
}

function loadDatasourceScreenWithSampleData(processedObject) {
    resetDatasourceScreen();
    if (selectedChart !== "SankeySentimentChart" && selectedChart !== "SankeyNonSentimentChart" && selectedChart !== "SankeySentimentChartAdvance" && selectedChart !== "SankeyNonSentimentChartAdvance") {
        $('#dropdownSheets').append($('<option value="' + 'Sample data' + '">' + 'Sample data' + '</option>'));
    }

    $('#chkHeaderRow').prop("checked", "checked");
    $("#txtBoxHeaderRow").val(1);
    $("#txtBoxHeaderRow").attr("disabled", "true");
    $("#chkHeaderRowImg").attr("src", "https://chartexpo.com/ChartExpoForGoogleSheetAddin/Images/v4/checkbox-active.png");
    //loading dimensions.
    if (processedObject != null && processedObject.headerRow.length > 0) {
        var chartObject = getSelectedChartRulesObject();
        dataSourceColumns[0] = processedObject.headerRow[0];
        dataSourceColumnsAll[0] = processedObject.headerRow[0];
        if (chartObject != null) {
            //loading searchable menu.
            loadOriginalColumns();
            //loadDimensions('dropDownDimensions0');
            var minimumRequiredDimensions = chartObject.MinDim;
            if (selectedChartNameFromSelectChartUI == "SankeySentimentChart" || selectedChartNameFromSelectChartUI == "SankeyNonSentimentChart") {
                minimumRequiredDimensions = 3;
            }
            else if (selectedChartNameFromSelectChartUI == "SankeySentimentChartAdvance" || selectedChartNameFromSelectChartUI == "SankeyNonSentimentChartAdvance") {
                minimumRequiredDimensions = 4;
            }
            for (var counter = 0; counter < minimumRequiredDimensions; counter++) {
                addDimensionsDropDownListForSampleData();
            }//end of for loop.

            for (var counter = 0; counter < chartObject.MinMetric; counter++) {
                addMeasuresDropDownListForSampleData();
            }

        }//end of inner if statement.

        //setting range control max value.            
        $('.slidecontainer').css('display', 'none');
        $('#selectedRowsDiv').css('display', 'none');
        $('#spanSelectRowRange').css('display', 'none');
        //$("#maxRangeValue").html(processedObject.dataRows.length);
    }//end of outer if statement.

    $('.dropdownchangeclass').prop("disabled", "true");
    $('.dimensionRemoveClass').css('visibility', 'hidden');
    $('.metricRemoveClass').css('visibility', 'hidden');
}

function openSelectedChartGuidInEditMode(chartGuid) {
    google.script.run.withSuccessHandler(function (addonInitialStateObj) {
        var myChartMeta = JSON.parse(addonInitialStateObj);

        $(".se-pre-con").fadeOut("slow");

        selectedChart = mySavedListSelectedChartName = myChartMeta.ChartName;
        var savedChartMetaJSON = myChartMeta.ChartMetaJSON;
        var savedHeaderRow = null, chartDimension = {};

        var useHeaderRow = "true";
        var a1NotationInformation = null;
        var dataRowFrom, dataRowTo, headerRowNumber, headerRowAnnotation, dataRowsAnnotation, selectedDimensions, selectedMeasures;

        mySavedListSelectedChartCategory = undefined;

        var editableMyChartCompleteDetail = {};

        // hanlde old saved my chart as need to show chart header count as per chart data
        mySavedListSelectedChartData = JSON.parse(myChartMeta.ChartDataJSON);

        /* no need to do it now, as in this case, all charts are drawn in latest format
        // get detail of the node to be 
        var chartAddedIntoMyChartListDate = new Date(+createdon); // createdOn

        if (chartAddedIntoMyChartListDate < googleSheetV3LaunchDate) {
            editableMyChartCompleteDetail.dataRows = getSampleDataRows(selectedChart, mySavedListSelectedChartData);
        }
        else {
            editableMyChartCompleteDetail.dataRows = mySavedListSelectedChartData;
        }*/

        editableMyChartCompleteDetail.dataRows = mySavedListSelectedChartData;
        //console.log("saved settings:" + JSON.stringify(savedChartMetaJSON));

        if (savedChartMetaJSON != undefined) {
            chartInEditModeSyncTime = true;
            savedChartMetaJSON = JSON.parse(savedChartMetaJSON); // parsed
            mySavedListSelectedChartCategory = savedChartMetaJSON.chartCategory;
            mySavedListSelectedChartProperties = savedChartMetaJSON.props; // parsed props
            chartDimension = savedChartMetaJSON.dimension;

            if (savedChartMetaJSON.headerRow != undefined) {
                savedHeaderRow = JSON.parse(savedChartMetaJSON.headerRow);
            }
            else {
                if (editableMyChartCompleteDetail.dataRows.length > 0) {
                    savedHeaderRow = getSampleDataHeaderRow(selectedChart, editableMyChartCompleteDetail.dataRows[0].length);
                }
                else {
                    savedHeaderRow = getSampleDataHeaderRow(selectedChart);
                }
            }

            if (mySavedListSelectedChartProperties != undefined) {
                mySavedListSelectedChartProperties = JSON.stringify(mySavedListSelectedChartProperties).replace("___br___", "<br>");
            }

            if (savedChartMetaJSON.selectedDimensions != undefined) {
                selectedDimensions = JSON.parse(savedChartMetaJSON.selectedDimensions);
            }
            if (savedChartMetaJSON.selectedMeasures != undefined) {
                selectedMeasures = JSON.parse(savedChartMetaJSON.selectedMeasures);
            }

            a1NotationInformation = savedChartMetaJSON.a1NotationInformation;


            //console.log("In edit chart start:" + savedChartMetaJSON.dataRowFrom);
            //console.log("In edit chart end:" + savedChartMetaJSON.dataRowTo);

            useHeaderRow = savedChartMetaJSON.useHeaderRow;
            fileName = savedChartMetaJSON.fileName;
            sheetName = savedChartMetaJSON.sheetName;
            fileId = savedChartMetaJSON.fileId;
            sheetId = savedChartMetaJSON.sheetId;
            headerRowAnnotation = savedChartMetaJSON.headerRowAnnotation;
            dataRowsAnnotation = savedChartMetaJSON.dataRowsAnnotation;
            headerRowNumber = savedChartMetaJSON.headerRowNumber;
            dataRowFrom = savedChartMetaJSON.dataRowFrom;
            dataRowTo = savedChartMetaJSON.dataRowTo;

            //console.log("use header row:" + useHeaderRow);
            //console.log("headerRowAnnotation:" + headerRowAnnotation);
            //console.log("sheetname:" + sheetName);
            //console.log("dataRowsAnnotation:" + dataRowsAnnotation);
            //console.log("fileName:" + fileName);
            //console.log("headerRowNumber:" + headerRowNumber);
            //console.log("startRowNumber:" + dataRowFrom);
            //console.log("endRowNumber:" + dataRowTo);
        }

        var chartRuleObject = getSelectedChartDisplayName();

        //decide show or hide Dimension/Measure Text
        isDimensionMeasureTextVisible(chartRuleObject);
        selectedchartDisplayName = chartRuleObject.ChartDisplayName;
        editableMyChartCompleteDetail.selectedChart = selectedChart;
        editableMyChartCompleteDetail.selectedChartDisplayName = selectedchartDisplayName;
        editableMyChartCompleteDetail.selectedChartCategory = mySavedListSelectedChartCategory;

        editableMyChartCompleteDetail.headerRow = savedHeaderRow;

        editableMyChartCompleteDetail.selectedDimensions = selectedDimensions;
        editableMyChartCompleteDetail.selectedMeasures = selectedMeasures;

        editableMyChartCompleteDetail.a1NotationInformation = a1NotationInformation;

        editableMyChartCompleteDetail.useHeaderRow = useHeaderRow;
        editableMyChartCompleteDetail.headerRowAnnotation = headerRowAnnotation;
        editableMyChartCompleteDetail.dataRowsAnnotation = dataRowsAnnotation;
        editableMyChartCompleteDetail.fileName = fileName;
        //alert("sheetName from db " + sheetName);
        editableMyChartCompleteDetail.sheetName = sheetName;
        editableMyChartCompleteDetail.fileId = savedChartMetaJSON.fileId;
        editableMyChartCompleteDetail.sheetId = savedChartMetaJSON.sheetId;
        editableMyChartCompleteDetail.headerRowNumber = headerRowNumber;
        editableMyChartCompleteDetail.dataRowFrom = dataRowFrom;
        editableMyChartCompleteDetail.dataRowTo = dataRowTo;

        editableMyChartCompleteDetail.defaultProperties = mySavedListSelectedChartProperties;
        synchedChartProperties = JSON.parse(mySavedListSelectedChartProperties); // used for next processing in case of sync mode
        synchedChartDimensions = chartDimension;

        // console.log("synchedChartProperties in edit chart " + synchedChartProperties);
        // console.log("synchedChartDimensions in edit chart " + synchedChartDimensions);

        editableMyChartCompleteDetail.myCharts = "true";
        editableMyChartCompleteDetail.dimension = chartDimension;
        editableMyChartCompleteDetail.editableChartGuid = chartGuid;

        //chartOpenInEditMode = true;
        onMyChartEditButtonClicked(editableMyChartCompleteDetail);
    })
        .withFailureHandler(
        function (msg, element) {
            $(".se-pre-con").fadeOut("slow");
            handleError(msg);
        }
        ).getSelectedChartMeta(chartGuid);
}
function isDimensionMeasureTextVisible(chartRuleObject) {
    //$("#divDimensionsContainer").css("margin-bottom", "0px");
    if (chartRuleObject.MaxDim > 1) {
        var dimension = chartRuleObject.DimensionText + "s";
        $("#dimensionText").html("");
        $("#dimensionText").html("Rearrange " + dimension.toLowerCase() + " with a drag-n-drop.");
        $("#dimensionText").show();
    }
    else
    {
        $("#dimensionText").hide();
    }
    if (chartRuleObject.MaxMetric > 1) {
        var metric = chartRuleObject.MeasureText + "s";
        $("#measureText").html("");
        $("#measureText").html("Rearrange " + metric.toLowerCase() + " with a drag-n-drop.");
        $("#measureText").show();
    }
    else {
        $("#measureText").hide();
    }
}

function addSampleSheetChartIntoMyChartList() {//headerColumnsArray, dataRowsAsSheet, newlyAddedSampleSheetDetail) {// fileName, sheetName, fileId, sheetId) {
    // add sample sheet data
    //1. select this sheet into dropdown list
    //2. set selected sheet name in template
    //3. set selectedDimensions
    //4. set selected measures
    //5. find rows count
    //6. set header row no
    //7. set proper a1NotationInformation
    //8. headerRow : "[[]]"
    //9. dataRowsAnnotation
    //10. headerRowAnnotation
    //11. sets chart default properties

    // headerColumnsArray contains : "[[\"Expense Type\",\"Subtopic_G1\",\"Sentiment_I1\",\"Amount\"]]"
    //var tblHeaderOfSampleSheet, tblBodyOfSampleSheet, newlyInsertedSampleSheetDetail;

    var headerColumnsArray = tblHeaderOfSampleSheet, dataRowsAsSheet = tblBodyOfSampleSheet, newlyAddedSampleSheetDetail = newlyInsertedSampleSheetCompleteDetail;

    var objectTemplate = {};
    selectedChartNameFromSelectChartUI = selectedChart;
    var chartRuleObject = getSelectedChartRulesObject();
    var mySavedChartCustomName = chartRuleObject.ChartDisplayName;
    var useHeaderRow = true;

    var newSampleChartDefaultProperties = [];
    if (selectedChart == "SankeySentimentChartAdvance" || selectedChart == "SankeyNonSentimentChartAdvance") {
        newSampleChartDefaultProperties = JSON.parse(chartDefaultProperties["SankeyNonSentimentChartAdvance"]);
    }
    else {
        newSampleChartDefaultProperties = getSelectedChartDefaultPropertiesSet();
    }

    var chartMetaJSON = {
        "props": newSampleChartDefaultProperties,
        "dimension": synchedChartDimensions,
        "chartCategory": selectedChartCategory,
        "headerRow": JSON.stringify(headerColumnsArray),
        "a1NotationInformation": '',
        "useHeaderRow": useHeaderRow,
        "dataRowsAnnotation": '',
        "headerRowAnnotation": '',
        "headerRowNumber": '',
        "dataRowFrom": '',
        "dataRowTo": '',
        "fileName": newlyAddedSampleSheetDetail.fileName,
        "sheetName": newlyAddedSampleSheetDetail.sheetName,
        "fileId": newlyAddedSampleSheetDetail.fileId,
        "sheetId": newlyAddedSampleSheetDetail.sheetId
    };

    var tempHeaderColumnsArray = JSON.parse(JSON.stringify(headerColumnsArray[0]));
    // alert("tempHeaderColumnsArray " + JSON.stringify(tempHeaderColumnsArray));
    var selectedDims = tempHeaderColumnsArray;
    var selectedMeasures = tempHeaderColumnsArray;

    //if (chartRuleObject == undefined ) {
    chartRuleObject = getSelectedChartRulesObject();
    //}
    if (tabularDataCharts_DimensionColumnName.hasOwnProperty(selectedChart)) {
        if (tabularDataCharts_DimensionColumnName[selectedChart].length == 1) {
            selectedMeasures = tempHeaderColumnsArray.splice(1);
            selectedDims = tempHeaderColumnsArray;
        }
        else if (tabularDataCharts_DimensionColumnName[selectedChart].length == 2) {
            selectedMeasures = tempHeaderColumnsArray.splice(2);
            selectedDims = tempHeaderColumnsArray;
        }

    }
    else if (selectedChart == "ParetoGroupedChart" || selectedChart == "ParetoGroupedHorizontalChart") {
        selectedMeasures = tempHeaderColumnsArray.splice(tempHeaderColumnsArray.length - chartRuleObject.MaxMetric, chartRuleObject.MaxMetric);
        selectedDims = tempHeaderColumnsArray;//.splice(0, tempHeaderColumnsArray.length - chartRuleObject.MinMetric);
    }
    else {
        // distribute part of selected header into dim and measures
        selectedMeasures = tempHeaderColumnsArray.splice(tempHeaderColumnsArray.length - chartRuleObject.MinMetric, chartRuleObject.MinMetric);
        selectedDims = tempHeaderColumnsArray;//.splice(0, tempHeaderColumnsArray.length - chartRuleObject.MinMetric);
    }

    // alert("selectedMeasures =" + JSON.stringify(selectedMeasures) + " selectedDims=" + JSON.stringify(selectedDims));
    var headerRowAnnotation = [];
    var dataRowsAnnotation = []; // contains data rows annotation information
    var a1NotationInformation = []; // it is array

    for (var col = 1; col < headerColumnsArray[0].length; col++) {
        a1NotationInformation.push("R2C" + col + ":" + "R" + dataRowsAsSheet.length + "C" + col);
        headerRowAnnotation.push("R1C" + col + ":" + "R1C" + col);
    }

    chartMetaJSON.selectedDimensions = JSON.stringify(selectedDims);
    chartMetaJSON.selectedMeasures = JSON.stringify(selectedMeasures);
    chartMetaJSON.a1NotationInformation = a1NotationInformation;
    chartMetaJSON.dataRowsAnnotation = JSON.stringify(a1NotationInformation);
    chartMetaJSON.headerRowAnnotation = JSON.stringify(headerRowAnnotation);
    chartMetaJSON.headerRowNumber = 1;
    chartMetaJSON.dataRowFrom = 0; //dataRowFrom;
    chartMetaJSON.dataRowTo = dataRowsAsSheet.length - 1; //dataRowTo;

    var returnedObj = getDatasourceScreenSampleChartState(newlyAddedSampleSheetDetail.sheetId, 1, selectedDims, selectedMeasures, dataRowsAsSheet.length,
        (selectedDims.length + selectedMeasures.length), headerColumnsArray, dataRowsAsSheet);

    // return;
    saveChartIntoMyChartsList(JSON.stringify(chartMetaJSON), selectedChart, mySavedChartCustomName, JSON.stringify(dataRowsAsSheet), "Add", chartGuid, JSON.stringify(returnedObj), true);

    // call send call to edit chart
}

// get chart properties
function getSelectedChartDefaultPropertiesSet() {
    var chartDefaultProperties = [];
    if (selectedChartCategory !== undefined && selectedChartCategory === "PPC") {
        if (PPCChartsDefaultProperties[selectedChart] != undefined) {
            chartDefaultProperties = PPCChartsDefaultProperties[selectedChart]();
            // alert("selectedChartCategory");
            //currentChartChangedPropertiesList = removeChartDefaultProperties(currentChartChangedPropertiesList, selectedChart, storedObjectInServerTempStorage["isSampleData"]);
        }
    }
    else {
        // else load PPC Charts properties
        if (DefaultProperties[selectedChart] != undefined) {
            // alert("selectedChartCategory not");
            chartDefaultProperties = DefaultProperties[selectedChart]();
            //currentChartChangedPropertiesList = removeChartDefaultProperties(currentChartChangedPropertiesList, selectedChart, storedObjectInServerTempStorage["isSampleData"]);
        }
    }

    if (chartDefaultProperties != undefined && chartDefaultProperties != null && chartDefaultProperties.length > 0) {
        // reformat Properties Format
        for (var i = 0; i < chartDefaultProperties.length; i++) {
            chartDefaultProperties[i]._guid = chartDefaultProperties[i]._guid + "_CMP_" + chartDefaultProperties[i]._componentType;
        }
    }

    var settings = {
        height: 420,
        width: 800, // new check added
        renderTo: "renderToDiv",
        data: [],
        lib: "D3",
        tooltipData: []
    };

    var chartObj = new ChartExpo[selectedChart](settings);

    if (chartDefaultProperties != null && chartDefaultProperties != undefined && Array.isArray(chartDefaultProperties) && chartDefaultProperties.length > 0) {
        // alert("updateDefaultPropertiesFromSavedDefaultProperties");
        var savedCurrentMergedPropertiesList = updateDefaultPropertiesFromSavedDefaultProperties(chartObj.getProperties(true), chartDefaultProperties);
        //currentChartChangedPropertiesList = savedProperties;
        chartObj = new ChartExpo[selectedChart](settings, savedCurrentMergedPropertiesList);
    }

    return chartObj.getProperties(true); // returned chart updated props
}

// set properties state from saved properties
function updateDefaultPropertiesFromSavedDefaultProperties(defaultProperties, savedProperties) {
    var actualPropsList = defaultProperties;
    for (var i = 0; i < actualPropsList.length; i++) {
        var newProp = {};
        newProp.propComp = actualPropsList[i].type;
        for (var j = 0; j < actualPropsList[i].childGroups.length; j++) {
            newProp.group = actualPropsList[i].childGroups[j].name;
            for (var k = 0; k < actualPropsList[i].childGroups[j].properties.length; k++) {
                actualPropsList[i].childGroups[j].properties[k]._value = getDefaultSavedPropertyValue(savedProperties, actualPropsList[i].childGroups[j].properties[k]._guid, actualPropsList[i].type, actualPropsList[i].childGroups[j].properties[k]._value);
            }
        }
    }
    return actualPropsList;
}

function getDefaultSavedPropertyValue(savedProperties, guid, componentName, unchangedValue) {
    for (var i = 0; i < savedProperties.length; i++) {
        var cmpName = savedProperties[i]._guid.split("_CMP_")[1];
        var propGuid = savedProperties[i]._guid.split("_CMP_")[0];

        if (propGuid == guid && cmpName == componentName) {
            return savedProperties[i]._value;
        }
    }
    return unchangedValue;
}

function getDatasourceScreenSampleChartState(sampleSheetId, headerRowNo, defaultDimensionsArray,
    defaultMetricsArray, sampleDataRowsCount, columnsCountInSampleData, headerColumnsTwoDimArray, dataRowsAsSheetFormat) {

    var drawChartObject = { ChartName: '', Sheet: '', HeaderRow: 0, Dimensions: [], Measures: [], RowStartIndex: 0, RowLastIndex: 0 };
    var dimensions = [];
    var measures = [];

    drawChartObject.ChartName = selectedChart;
    drawChartObject.Sheet = sampleSheetId;//$('#dropdownSheets').val();
    drawChartObject.HeaderRow = headerRowNo; // headerRowNo
    dimensions = defaultDimensionsArray;

    drawChartObject.Dimensions = defaultDimensionsArray;
    drawChartObject.Measures = defaultMetricsArray;
    drawChartObject.RowLastIndex = sampleDataRowsCount;

    var rowFrom = 1;
    var rowTo = sampleDataRowsCount - 1;
    var oldSelectedCellsA1Notations = [];

    var headerRowAnnotation = [];
    var dataRowsAnnotation = []; // contains data rows annotation information
    var a1NotationInformation = []; // it is array

    for (var col = 1; col < headerColumnsTwoDimArray[0].length; col++) {
        a1NotationInformation.push("R2C" + col + ":" + "R" + dataRowsAsSheetFormat.length + "C" + col);
        headerRowAnnotation.push("R1C" + col + ":" + "R1C" + col);
    }

    // while dataRowsAnnotation is string
    var finalState = {
        useHeaderRow: "", fileName: "", fileId: "", sheetName: drawChartObject.Sheet, sheetId: "",
        headerRow: headerColumnsTwoDimArray, dataRows: dataRowsAsSheetFormat, a1NotationInformation: a1NotationInformation,
        "headerRowNumber": headerRowNo, "dataRowFrom": rowFrom, "dataRowTo": rowTo,
        "headerRowAnnotation": JSON.stringify(headerRowAnnotation), "dataRowsAnnotation": JSON.stringify(a1NotationInformation),
        "selectedDimensions": JSON.stringify(drawChartObject.Dimensions), "selectedMeasures": JSON.stringify(drawChartObject.Measures)
    };
    //alert("on save " + $("#txtBoxHeaderRow").val());
    //console.log("on save " + $("#txtBoxHeaderRow").val());

    return finalState;
}
function getValueInKFormat(number) {
    var numberFormat;
    number = (+number);
    if (number > 0 && number < 10) {
        numberFormat = d3.format(".1s");
    }
    else if (number > 1000) {
        numberFormat = d3.format(".3s");
    }
    else {
        numberFormat = d3.format(".2s");
    }

    return numberFormat(number);
}
function showReconnectingOverlay() {
    $(".reconnecting-overlay").show();
}
function hideReconnectingOverlay() {
    $(".reconnecting-overlay").hide();
}
var checkInternetHandler = null;
function checkInternetConnection() {
    if (showInternetReconnectingOverlay) {
        var timeOut;
        var ifConnected = window.navigator.onLine;
        if (ifConnected) {
            hideReconnectingOverlay();
            timeOut = setTimeout(checkInternetHandler, 2000);
            clearTimeout(timeOut);
        }
        return ifConnected;
    }
    else {
        return true;
    }
    // if internet come
    // hide overlay
    // set timeout of timeline
    // clearTimeout(checkInternetHandler);
}

function handleError(msg) {
    if (showInternetReconnectingOverlay) {
        if (msg == "NetworkError: Connection failure due to HTTP 0") {
            $(".se-pre-con").fadeOut("slow");
            showReconnectingOverlay();
            checkInternetHandler = setInterval(checkInternetConnection, 3000);
        }
    }
    else {
        alert(msg);
    }
    //1. check what message shown in case of internet disconnectivity
    //if message contains any net related message{
    // show overlay 
    // start timer to check internet after 
    //checkInternetHandler = setInterval(checkConnection, 3000);
}

function clearMeasureForAllCurrencies(number) {
    // handle true & false as well
    if (number == null || number == undefined || number == "") {
        return 0;
    }
    else if (number.toLowerCase() == "true") {//If measure contains boolean value.
        return 1;
    }
    else if (number.toLowerCase() == "false") {//If measure contains boolean value.
        return 0;
    }

    var remAllAlphabetsReg = /[^0-9.,-]/g;  // th
    //var number = 'fr-12, 34-56. 25- dd $';
    //var number = 'fr-125,2547,458,265 dd $';

    //var number = '1.234,56 ₫';

    var numberWithoutAlphabetsAndSymbols = number.replace(remAllAlphabetsReg, '');
    var removeAllDashedRegExp = /[^0-9.,]/g;
    var withoutDashed = numberWithoutAlphabetsAndSymbols.replace(removeAllDashedRegExp, '');

    var addNegativeDash = false;
    if (numberWithoutAlphabetsAndSymbols.length > 0 && numberWithoutAlphabetsAndSymbols[0] == '-') {
        addNegativeDash = true;
    }

    //alert("withoutDashed=" + withoutDashed);
    if (withoutDashed.length == 0) {
        return 0;
    }

    // if
    while (withoutDashed[withoutDashed.length - 1] == '.') {
        withoutDashed = withoutDashed.substring(0, withoutDashed.length - 2);
    }
    while (withoutDashed[0] == '.') {
        withoutDashed = withoutDashed.substring(1);
    }

    // remove all dots if appeared before or at the end of 
    // to do
    // fr. 1.234,56 // before
    // 

    // decide about thousands separator and replace 
    var dotIndexes = [], commaIndexes = [];
    for (var i = 0; i < withoutDashed.length; i++) {
        if (withoutDashed[i] == ',') {
            commaIndexes.push(i);
        }
        if (withoutDashed[i] == '.') {
            dotIndexes.push(i);
        }
    }

    //alert("dotIndexes= "+JSON.stringify(dotIndexes) + " commaIndexes= "+JSON.stringify(commaIndexes));
    var finalNumberAfterCleansing = withoutDashed;
    if (commaIndexes.length > 0 && dotIndexes.length == 0) {
        // 125,2547,458
        // only comma exists, check is it thousands separator or digits separator
        if (commaIndexes.length > 1) {
            // its mean it is thousands separator
            var removeAllCommasExp = /[^0-9]/g;
            finalNumberAfterCleansing = withoutDashed.replace(removeAllCommasExp, '');
        }
        else if (commaIndexes.length == 1) {
            // check is it thousands separator or digits separator
            // 15,25 - 4
            // 152,236
            // 1,225000
            // find number of digits before and after comma
            var l = withoutDashed.length;
            var afterCommaDigitsLength = withoutDashed.length - (+commaIndexes[0]);
            var beforeCommaDigitsLength = withoutDashed.length - (+commaIndexes[0]);
            //debugger;
            if (afterCommaDigitsLength < 4 || afterCommaDigitsLength > 4) {
                //digit separater
                finalNumberAfterCleansing = withoutDashed.replace(',', '.');
            }
            else if (afterCommaDigitsLength == 4 && beforeCommaDigitsLength < 4) {
                // thousands separator         
                finalNumberAfterCleansing = withoutDashed.replace(',', '.');
            }
            else {
                finalNumberAfterCleansing = withoutDashed.replace(',', '');
            }
        }
    }
    else if (commaIndexes.length == 0 && dotIndexes.length > 0) {
        // 125,2547,458
        // only comma exists, check is it thousands separator or digits separator
        if (dotIndexes.length > 1) {
            // its mean it is thousands separator
            var removeAllCommasExp = /[^0-9]/g;
            finalNumberAfterCleansing = withoutDashed.replace(removeAllCommasExp, '');
        }
        else if (dotIndexes.length == 1) {
            // check is it thousands separator or digits separator
            // 15,25 - 4
            // 152,236
            // 1,225000
            // find number of digits before and after comma
            var l = withoutDashed.length;
            var afterCommaDigitsLength = withoutDashed.length - (+dotIndexes[0]);
            var beforeCommaDigitsLength = withoutDashed.length - (+dotIndexes[0]);

            if (afterCommaDigitsLength < 4 || afterCommaDigitsLength > 4) {
                //digit separater
                // do nothing as it is already correct
                //finalNumberAfterCleansing = withoutDashed.replace(',', '.');
            }
            else if (afterCommaDigitsLength == 4 && beforeCommaDigitsLength < 4) {
                // thousands separator
                finalNumberAfterCleansing = withoutDashed.replace('.', '');
            }
            else {
                // retain it as it is
                finalNumberAfterCleansing = withoutDashed.replace('.', '.');
            }
        }
    }
    else if (commaIndexes.length > 0 && dotIndexes.length > 0) {
        // 125,254,236.25

        //  find which one come (, or .) appeared at the end
        var lastOccurrenceCharacter = '';
        lastOccurrenceCharacter = +commaIndexes[commaIndexes.length - 1] > +dotIndexes[dotIndexes.length - 1] ? withoutDashed[+commaIndexes[commaIndexes.length - 1]] : withoutDashed[+dotIndexes[dotIndexes.length - 1]]; //alert("ddd " + lastOccurrenceCharacter);
        if (lastOccurrenceCharacter == '.') {
            // remove all commas and treat them as thousands separator
            // 
            var removeAllCommasRegExp = /[^0-9.]/g;
            finalNumberAfterCleansing = withoutDashed.replace(removeAllCommasRegExp, '');
        }
        else if (lastOccurrenceCharacter == ',') {
            // 125.125.236,25
            // remove all dots and treat them as thousands separator and replace last comma with dot to treat like digits
            var removeAllCommasRegExp = /[^0-9,]/g;
            finalNumberAfterCleansing = withoutDashed.replace(removeAllCommasRegExp, '');
            //alert("ddd " + finalNumberAfterCleansing);
            finalNumberAfterCleansing = finalNumberAfterCleansing.replace(',', '.');
        }
    }

    if (isNaN(finalNumberAfterCleansing)) {
        //alert("0");
        return 0;
    }

    if (addNegativeDash) {
        var finalNumber = '-' + finalNumberAfterCleansing;
        //alert("with dash=" + finalNumber);
    }

    return finalNumberAfterCleansing;
}
//</script >