/// <reference path="jquery.js" />

///// constants and global variables

var const_showAllLinks;
var const_textIfNoDetailedNotes;
var const_importantVerseWeight;

var global_IsDetailed;
var global_TrackbarValue;

///// initialization

$(function () {
    bindEvents();
});

function setConstants(showAllLinks, textIfNoDetailedNotes, importantVerseWeight) {
    const_showAllLinks = showAllLinks;
    const_textIfNoDetailedNotes = textIfNoDetailedNotes;
    const_importantVerseWeight = importantVerseWeight;
}

function initFilter(notebooks, trackbarValue, showDetailedNotes) {
    global_IsDetailed = showDetailedNotes;
    global_TrackbarValue = trackbarValue;

    setFilterNotebooks(jQuery.parseJSON(notebooks));
    setFilterLinkHandlers();
    setDetailedNotesFilter();
    setFilterTrackbar();
    setFilterCommonHandlers();
    setFilterChangedHandler();

    $(".saveFilterSettings").hide();
}

///// filter methods

function setFilterNotebooks(notebooks) {
    var table = $("#notebooksFilterTable");
    for (var i = 0; i < notebooks.length; i++) {
        var id = "filterNotebook_" + i;
        var syncId = notebooks[i].SyncId;
        var title = notebooks[i].Title;
        var checked = notebooks[i].Checked ? "checked" : "";
        table.append("<tr><td class='notebooksFilter'><input type='checkbox' class='notebooksFilter' id='" + id + "' syncId='" + syncId + "' " + checked + " /><label for='" + id + "' class='notebooksFilter'>" + title + "</label></td></tr>");

        filterNotebook(syncId, checked);
    }
}

function setDetailedNotesFilter() {
    var chkDetailedNodes = $("#chkDetailedNotes");
    chkDetailedNodes.click(chkDetailedNodes_click);
    chkDetailedNodes.attr('checked', global_IsDetailed);

    var detailedEls = $(".detailed");
    if (detailedEls.length == 0) {
        chkDetailedNodes.attr("disabled", "disabled");

        chkDetailedNodes.attr("title", const_textIfNoDetailedNotes);
        $("label.detailedNotes").attr("title", const_textIfNoDetailedNotes);
    }
}

function setFilterLinkHandlers() {
    $("#filterLink").click(function (e) {
        e.preventDefault();
        e.stopPropagation();
        $('#filter').fadeIn(100).focus();
    });

    $("[class*=filterPopup]").click(function (e) {
        e.stopPropagation();        // Предотвращаем работу ссылки, если она являеться нашим popup окном 
    });

    $("html").click(function () {
        var scrollPos = $(window).scrollTop();
        hideFilter();
    });
}

function setFilterChangedHandler() {
    $("input.notebooksFilter").click(notebookFilterClick);
    $("#saveFilterSettingsLink").click(saveFilterSettings);
}

var linksWeight;
var firstTrackbarAndDetailedNotesFilterCall = true;
function filterByTrackbarAndDetailedNotes() {
    if (!(firstTrackbarAndDetailedNotesFilterCall && global_TrackbarValue == 1 && global_IsDetailed)) {           // иначе нам ничего не надо скрывать

        if (linksWeight == null)
            linksWeight = BuildLinksWeight();

        for (var index = 0; index < linksWeight.length; index++) {
            var linkInfo = linksWeight[index];

            var weight = global_IsDetailed ? linkInfo.detailedWeight : linkInfo.weight;

            setElementIsImportant(linkInfo.link, weight >= const_importantVerseWeight);

            hideLinkIfIsNotDetailed(linkInfo);

            var isAnyChildVisible = false;
            for (var childIndex = 0; childIndex < linkInfo.childLinks.length; childIndex++) {
                var childLinkInfo = linkInfo.childLinks[childIndex];
                hideLinkIfIsNotDetailed(childLinkInfo);
                isAnyChildVisible = hideLinkIfLessThanTrackbarValue(childLinkInfo.link, childLinkInfo.weight < global_TrackbarValue) || isAnyChildVisible;
            }

            if (linkInfo.childLinks.length == 0)
                hideLinkIfLessThanTrackbarValue(linkInfo.link, weight < global_TrackbarValue);
            else
                hideLinkIfLessThanTrackbarValue(linkInfo.link, !isAnyChildVisible);
        }
    }

    firstTrackbarAndDetailedNotesFilterCall = false;
}

function setElementIsImportant(linkEl, isImportant) {
    if (isImportant)
        linkEl.addClass("importantVerseLink");
    else
        linkEl.removeClass("importantVerseLink");
}

function hideLinkIfLessThanTrackbarValue(linkEl, toHide) {
    if (toHide) {
        hideElement(linkEl, "hiddenByTrackbarFilter");
        return false;
    }
    else {
        showElement(linkEl, "hiddenByTrackbarFilter");
        return true;
    }
}

function hideLinkIfIsNotDetailed(linkInfo) {
    if (linkInfo.isDetailed) {
        if (!global_IsDetailed)
            hideElement(linkInfo.link, "hiddenDetailed");
        else
            showElement(linkInfo.link, "hiddenDetailed");
    }
}

function BuildLinksWeight() {
    var titleLinks = [];
    $(".levelTitleLink").each(function (index) {
        var titleLinkEl = $(this);
        var isDetailed = titleLinkEl.hasClass("detailed");
        var weight = parseFloat(getParameter(this, "vw"));
        var detailedWeight = weight;
        var childLinks = [];

        if (isNaN(weight)) {
            titleLinkEl.parents(".pageLevel").find(".subLinkLink").each(function (childIndex) {
                var childLinkWeight = parseFloat(getParameter(this, "vw"));
                var childLinkEl = $(this);
                var childLinkIsDetailed = childLinkEl.hasClass("detailed");

                if (childLinkIsDetailed)
                    detailedWeight = isNull(detailedWeight, 0) + childLinkWeight;
                else
                    weight = isNull(weight, 0) + childLinkWeight;

                childLinks.push({ link: childLinkEl, isDetailed: childLinkIsDetailed, weight: childLinkWeight });
            });
        }

        titleLinks.push({ link: titleLinkEl, isDetailed: isDetailed, weight: weight, detailedWeight: detailedWeight, childLinks: childLinks });
    });

    return titleLinks;
}

function isNull(v, defaultValue) {
    if (isNaN(v) || v == undefined || v == null)
        return defaultValue;
    else return v;
}

function getParameter(str, name) {
    if (name = (new RegExp('[?&]' + encodeURIComponent(name) + '=([^&]*)')).exec(str))
        return decodeURIComponent(name[1]);
}

function notebookFilterClick() {
    var sender = $(this);
    filterNotebook(sender.attr("syncId"), sender.is(':checked'));

    filterWasChanged();
}

function filterNotebook(syncId, checked) {
    var notebookEl = $(".notebookLevel[syncid='" + syncId + "']");
    if (checked)
        showElement(notebookEl, "hiddenNotebook");
    else
        hideElement(notebookEl, "hiddenNotebook");
}

function chkDetailedNodes_click() {
    var chkDetailedNodes = $("#chkDetailedNotes");
    global_IsDetailed = chkDetailedNodes.is(":checked");

    filterByTrackbarAndDetailedNotes();

    //window.external.chkDetailedNodes_Changed(global_IsDetailed);    

    //document.location = document.location;   // если будут проблемы с обновлением нумерации у списков, раскомментировать эту линию 

    filterWasChanged();
}

function saveFilterSettings() {
    alert("saved");
    hideFilter();
}

var hideFilterTimerId;
function setFilterCommonHandlers() {
    $("#filter").mouseleave(function () { hideFilterTimerId = window.setTimeout(function () { hideFilter(); }, 1000); });
    $("#filter").mouseenter(function () { clearTimeout(hideFilterTimerId); });
}

function hideFilter() {
    $("#filter").fadeOut(100);
}

function filterWasChanged() {
    $(".saveFilterSettings").fadeIn();
}

function setFilterTrackbarTitle() {
    $("#filterVerseWeightTitle").html(global_TrackbarValue == 0 ? const_showAllLinks : global_TrackbarValue);
    if (global_TrackbarValue == 0)
        $("#filterVerseWeightDescription").hide();
    else
        $("#filterVerseWeightDescription").show();
}

function getFilterTrackbarValue(trackbarValue) {
    var values = [0, 0.1, 0.2, 0.25, 0.3, 0.5, 1, 2];
    return values[trackbarValue];
}

function setFilterTrackbar() {
    if ($("#filterVerseWeight").length > 0) {
        trackbar.getObject('chars').init({
            onMove: function () {
                global_TrackbarValue = getFilterTrackbarValue(this.leftValue);
                setFilterTrackbarTitle();
                filterByTrackbarAndDetailedNotes();
                filterWasChanged();
            },
            dual: false,
            width: 100, // px
            leftLimit: 0, // unit of value        
            rightLimit: 7, // unit of value
            leftValue: global_TrackbarValue, // unit of value
            clearLimits: true,
            clearValues: true
        },
        'filterVerseWeight');
    }
}

function hideElement(el, className) {
    if (el.length > 0 && !el.hasClass(className)) {
        var nodeName = el[0].nodeName;
        var parentEl = $(el[0].parentNode);
        el.addClass(className)

        if (!el.hasClass("verseLevel") && !el.hasClass("subLinkDelimeter") && !el.hasClass("subLinkMultiVerse")) {
            if (parentEl.children(nodeName).length == parentEl.children(nodeName + "." + className).length)
                hideElement(parentEl, className);
        }
    }
}

function showElement(el, className) {
    if (el.length > 0 && el.hasClass(className)) {
        var nodeName = el[0].nodeName;
        var parentEl = $(el[0].parentNode);

        if (!el.hasClass("verseLevel") && !el.hasClass("subLinkDelimeter") && !el.hasClass("subLinkMultiVerse")) {
            if (parentEl.children(nodeName).length >= parentEl.children(nodeName + "." + className).length)
                showElement(parentEl, className);
        }

        el.removeClass(className)
    }
}




///// helper methods

function bindEvents() {
    $("img.minus").mouseenter(minus_mouseIn).mouseleave(minus_mouseOut).click(minus_click);
    $("span.levelTitleText, a.verseLink").mouseenter(level_mouseIn).mouseleave(level_mouseOut);
}

function minus_click(event) {
    var target = $(event.target);
    if (!target.hasClass("collapsed")) {
        target.addClass("collapsed");
        target.attr("src", "../../../images/plus.png");
        target.parent().children("ol.levelChilds").addClass("collapsedLevel");
    }
    else {
        target.removeClass("collapsed");
        target.attr("src", "../../../images/minus.png");
        target.parent().children("ol.levelChilds").removeClass("collapsedLevel");
    }
}

function minus_mouseIn(event) {
    var target = $(event.target);
    processMinusImgIn(target);
}

function minus_mouseOut(event) {
    var target = $(event.target);
    processMinusImgOut(target);
}

function level_mouseIn(event) {
    var target = $(event.target);
    var img = $(target[0].parentNode.parentNode).children("img.minus");
    processMinusImgIn(img);
}

function level_mouseOut(event) {
    var target = $(event.target);
    var img = $(target[0].parentNode.parentNode).children("img.minus");
    processMinusImgOut(img);
}

function processMinusImgIn(img) {
    if (!img.hasClass("collapsed"))
        img.attr("src", "../../../images/minus.png");
}

function processMinusImgOut(img) {
    if (!img.hasClass("collapsed"))
        img.attr("src", "../../../images/none.png");
}