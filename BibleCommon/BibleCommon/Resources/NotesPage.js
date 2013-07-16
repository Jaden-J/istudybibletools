/// <reference path="jquery.js" />

///// constants and global variables

var const_showAllLinks;
var const_textIfNoDetailedNotes;
var const_importantVerseWeight;

var global_IsDetailed;
var global_MinVersesWeight;

///// initialization

$(function () {
    bindEvents();
});

function setConstants(showAllLinks, textIfNoDetailedNotes, importantVerseWeight) {
    const_showAllLinks = showAllLinks;
    const_textIfNoDetailedNotes = textIfNoDetailedNotes;
    const_importantVerseWeight = importantVerseWeight;
}

function initFilter(notebooks, minVersesWeight, showDetailedNotes) {
    global_IsDetailed = showDetailedNotes;
    global_MinVersesWeight = minVersesWeight;

    setFilterNotebooks(jQuery.parseJSON(notebooks));
    setFilterLinkHandlers();
    setDetailedNotesFilter();
    setFilterTrackbar();
    setFilterCommonHandlers();
    setFilterChangedHandler();

    hideFilterSettings();    
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

    //    var detailedEls = $(".detailed");
    //    if (detailedEls.length == 0) {
    //        chkDetailedNodes.attr("disabled", "disabled");

    //        chkDetailedNodes.attr("title", const_textIfNoDetailedNotes);
    //        $("label.detailedNotes").attr("title", const_textIfNoDetailedNotes);
    //    }
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
    if (!(firstTrackbarAndDetailedNotesFilterCall && global_MinVersesWeight == 0 && global_IsDetailed)) {           // иначе нам ничего не надо скрывать

        if (linksWeight == null)
            linksWeight = BuildLinksWeight();

        for (var index = 0; index < linksWeight.length; index++) {
            var linkInfo = linksWeight[index];
            var childsWeight = 0;

            linkInfo.visible = hideLinkIfIsNotDetailed(linkInfo.linkEl, linkInfo.isDetailed);

            var visibleChildrenCount = 0;
            for (var childIndex = 0; childIndex < linkInfo.childLinks.length; childIndex++) {
                var childLinkInfo = linkInfo.childLinks[childIndex];
                childLinkInfo.visible = hideLinkIfIsNotDetailed(childLinkInfo.linkEl, childLinkInfo.isDetailed);
                childLinkInfo.visible = hideLinkIfLessThanTrackbarValue(childLinkInfo.linkEl, childLinkInfo.weight < global_MinVersesWeight) && childLinkInfo.visible;

                if (childLinkInfo.visible) {
                    childsWeight += childLinkInfo.weight;
                    visibleChildrenCount++;
                }

                if (!isNull(childLinkInfo.nextBracketsEl)) {
                    hideLinkIfIsNotDetailed(childLinkInfo.nextBracketsEl, childLinkInfo.isDetailed);
                    hideLinkIfLessThanTrackbarValue(childLinkInfo.nextBracketsEl, childLinkInfo.weight < global_MinVersesWeight);
                }
                if (!isNull(childLinkInfo.nextDelimiterEl)) {
                    hideLinkIfIsNotDetailed(childLinkInfo.nextDelimiterEl, childLinkInfo.isDetailed);
                    hideLinkIfLessThanTrackbarValue(childLinkInfo.nextDelimiterEl, childLinkInfo.weight < global_MinVersesWeight);
                }
            }

            var weight = global_IsDetailed ? linkInfo.detailedWeight : linkInfo.weight;

            if (linkInfo.childLinks.length == 0) {
                linkInfo.visible = hideLinkIfLessThanTrackbarValue(linkInfo.linkEl, weight < global_MinVersesWeight) && linkInfo.visible;
                setElementIsImportant(linkInfo.linkEl, weight >= const_importantVerseWeight);
            }
            else {
                linkInfo.visible = hideLinkIfLessThanTrackbarValue(linkInfo.linkEl, visibleChildrenCount == 0) && linkInfo.visible;
                setElementIsImportant(linkInfo.linkEl, childsWeight >= const_importantVerseWeight);
            }

            if (linkInfo.visible) {
                if (visibleChildrenCount == 1) {
                    copyChildLinkToPageTitleLink(linkInfo);
                }
                else {
                    resetPageTitleLink(linkInfo);
                }
            }
        }
    }

    firstTrackbarAndDetailedNotesFilterCall = false;
}

function copyChildLinkToPageTitleLink(linkInfo) {
    var visibleChildLinkInfo;
    for (var i = 0; i < linkInfo.childLinks.length; i++) {
        if (linkInfo.childLinks[i].visible) {
            visibleChildLinkInfo = linkInfo.childLinks[i];
            break;
        }
    }
    if (visibleChildLinkInfo != null) {
        var multiVerseString;
        var childHref = visibleChildLinkInfo.linkEl.attr("href");
        if (!isNull(visibleChildLinkInfo.nextBracketsEl)) {
            multiVerseString = visibleChildLinkInfo.nextBracketsEl.text();
        }

        if (isNull(linkInfo.linkEl.attr("data_oldHref"))) {

            linkInfo.linkEl.attr("data_oldHref", linkInfo.linkEl.attr("href"));
            linkInfo.linkEl.attr("href", childHref);
            if (!isNull(multiVerseString)) {
                linkInfo.linkEl.after("<span class='childSubLinkMultiVerse'> " + multiVerseString + "</span>");
            }

            visibleChildLinkInfo.linkEl.addClass("hiddenChildLink");

            if (!isNull(visibleChildLinkInfo.nextBracketsEl)) {
                visibleChildLinkInfo.nextBracketsEl.addClass("hiddenChildLink");
            }

            if (!isNull(visibleChildLinkInfo.nextDelimiterEl)) {
                visibleChildLinkInfo.nextDelimiterEl.addClass("hiddenChildLink");
            }
        }
    }
}

function resetPageTitleLink(linkInfo) {
    if (!isNull(linkInfo.linkEl.attr("data_oldHref"))) {
        linkInfo.linkEl.attr("href", linkInfo.linkEl.attr("data_oldHref"));
        linkInfo.linkEl.attr("data_oldHref", null)
    }
    var parentEl = linkInfo.linkEl.parents("li")
    parentEl.find(".childSubLinkMultiVerse").remove();
    parentEl.find(".hiddenChildLink").removeClass("hiddenChildLink");
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

function hideLinkIfIsNotDetailed(linkEl, isLinkDetailed) {
    if (isLinkDetailed) {
        if (!global_IsDetailed) {
            hideElement(linkEl, "hiddenDetailed");
            return false;
        }
        else {
            showElement(linkEl, "hiddenDetailed");
            return true;
        }
    }
    return true;
}

function BuildLinksWeight() {
    var titleLinks = [];
    $(".levelTitleLink").each(function (index) {
        var titleLinkEl = $(this);
        var isDetailed = titleLinkEl.hasClass("detailed");
        var weight = parseFloat(isNullOrDefault(getParameter(this, "vw"), "0").replace(',', '.'));
        var detailedWeight = weight;
        var childLinks = [];

        if (weight == "0") {
            titleLinkEl.parents(".pageLevel").find(".subLinkLink").each(function (childIndex) {
                var childLinkWeight = parseFloat(getParameter(this, "vw").replace(',', '.'));
                var childLinkEl = $(this);
                var childLinkIsDetailed = childLinkEl.hasClass("detailed");

                if (!childLinkIsDetailed)
                    weight = isNullOrDefault(weight, 0) + childLinkWeight;

                detailedWeight = isNullOrDefault(detailedWeight, 0) + childLinkWeight;

                var parentTdEl = childLinkEl.parent();
                var nextBracketsEl = parentTdEl.nextAll(".subLinkMultiVerse").first();
                var nextDelimiterEl = parentTdEl.nextAll(".subLinkDelimeter").first();

                if (nextBracketsEl.length > 0 || nextDelimiterEl.length > 0) {
                    var nextVerseEl = parentTdEl.nextAll(".subLink").first();
                    if (nextVerseEl.length > 0) {
                        var nextVerseIndex = nextVerseEl.index();
                        if (nextVerseIndex < nextBracketsEl.index())
                            nextBracketsEl = null;
                        if (nextVerseIndex < nextDelimiterEl.index())
                            nextDelimiterEl = null;
                    }
                }

                childLinks.push({ linkEl: childLinkEl, nextBracketsEl: nextBracketsEl, nextDelimiterEl: nextDelimiterEl, isDetailed: childLinkIsDetailed, weight: childLinkWeight, visible: true });
            });
        }

        titleLinks.push({ linkEl: titleLinkEl, isDetailed: isDetailed, weight: weight, detailedWeight: detailedWeight, childLinks: childLinks, visible: true });
    });

    return titleLinks;
}

function isNullOrDefault(v, defaultValue) {
    if (isNull(v))
        return defaultValue;
    else
        return v;
}

function isNull(v) {
    return (v == undefined || v == null);
}

function getParameter(str, name) {
    if (name = (new RegExp("[?&]" + encodeURIComponent(name) + "=([^&]*)")).exec(str))
        return decodeURIComponent(name[1]);
}

function notebookFilterClick() {
    var sender = $(this);
    filterNotebook(sender.attr("syncId"), sender.is(":checked"));

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

    //document.location = document.location;   // если будут проблемы с обновлением нумерации у списков, раскомментировать эту линию 

    filterWasChanged();
}

function saveFilterSettings() {
    window.external.SaveFilterSettings(getHiddenNotebooks(), global_MinVersesWeight, global_IsDetailed.toString());
    hideFilter();
    hideFilterSettings();
}


function hideFilterSettings() {
    $(".saveFilterSettings").hide();
}


function getHiddenNotebooks() {
    var result = "";

    $("input.notebooksFilter").each(function (index) {
        var notebookEl = $(this);
        if (!notebookEl.is(":checked")) {
            result += notebookEl.attr("syncid") + "_|_";
        }
    });

    return result;
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
    $("#filterVerseWeightTitle").html(global_MinVersesWeight == 0 ? const_showAllLinks : global_MinVersesWeight);
    if (global_MinVersesWeight == 0)
        $("#filterVerseWeightDescription").hide();
    else
        $("#filterVerseWeightDescription").show();
}

function getMinVersesWeight(trackbarValue) {
    var values = [0, 1 / 10, 1 / 5, 1 / 4, 3 / 10, 1 / 2, 1, 2];
    return values[trackbarValue];
}

function getFilterTrackbarValue(minVersesWeight) {
    var values = [0, 1 / 10, 1 / 5, 1 / 4, 3 / 10, 1 / 2, 1, 2];

    for (var i = 0; i < values.length; i++)
        if (values[i] == minVersesWeight)
            return i;

    return 0;
}

function setFilterTrackbar() {
    if ($("#filterVerseWeight").length > 0) {
        trackbar.getObject('chars').init({
            onMove: function () {
                global_MinVersesWeight = getMinVersesWeight(this.leftValue);
                setFilterTrackbarTitle();
                filterByTrackbarAndDetailedNotes();
                filterWasChanged();
            },
            dual: false,
            width: 100, // px
            leftLimit: 0, // unit of value        
            rightLimit: 7, // unit of value
            leftValue: getFilterTrackbarValue(global_MinVersesWeight), // unit of value
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