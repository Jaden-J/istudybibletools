$(function () {
    bindEvents();
});

function bindEvents() {
    var chkDetailedNodes = $("#chkDetailedNotes");
    chkDetailedNodes.click(chkDetailedNodes_click);

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

function initDetailedNotes(showDetailedNotes, textIfNoDetailedNotes) {
    var chkDetailedNodes = $("#chkDetailedNotes");
    chkDetailedNodes.attr('checked', showDetailedNotes);

    var detailedEls = $(".detailed");
    if (detailedEls.length == 0) {
        chkDetailedNodes.attr("disabled", "disabled");

        chkDetailedNodes.attr("title", textIfNoDetailedNotes);
        $("label.detailedNotes").attr("title", textIfNoDetailedNotes);
    }
    else if (!showDetailedNotes) {
        hideDetailedElements(detailedEls);
    }
}

function chkDetailedNodes_click() {
    var chkDetailedNodes = $("#chkDetailedNotes");
    var checked = chkDetailedNodes.is(":checked");

    window.external.chkDetailedNodes_Changed(checked);

    //document.location = document.location;   // если будут проблемы с обновлением нумерации у списков, раскомментировать эту линию и закомментировать последующие.
    
    if (!checked) {
        var detailedEls = $(".detailed");
        hideDetailedElements(detailedEls);
    }
    else {
        var hiddenElemenents = $(".hiddenDetailed");
        hiddenElemenents.removeClass("hiddenDetailed");
    }    
}

function hideDetailedElements(elements) {
    for (var i = 0; i < elements.length; i++) {
        var el = elements[i];
        hideDetailedElement(el);
    }
}

function hideDetailedElement(el) {

    var $el = $(el);
    var $parent = $(el.parentNode);    

    $el.addClass("hiddenDetailed")

    if (!$el.hasClass("verseLevel") && !$el.hasClass("subLinkDelimeter") && !$el.hasClass("subLinkMultiVerse")) {

        if ($parent.children(el.nodeName).length == $parent.children(el.nodeName + ".hiddenDetailed").length)
            hideDetailedElement(el.parentNode);
    }
}