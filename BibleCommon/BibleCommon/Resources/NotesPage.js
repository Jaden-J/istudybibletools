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
    var $target = $(event.target);
    if (!$target.hasClass("collapsed")) {
        $target.addClass("collapsed");
        $target.attr("src", "../../../images/plus.png");        
        $target.parent().children("ol.levelChilds").addClass("hidden");
    }
    else {
        $target.removeClass("collapsed");
        $target.attr("src", "../../../images/minus.png");
        $target.parent().children("ol.levelChilds").removeClass("hidden");
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

function initDetailedNotes(showDetailedNotes) {
    var chkDetailedNodes = $("#chkDetailedNotes");    
    chkDetailedNodes.attr('checked', showDetailedNotes);

    var detailedEls = $(".detailed");
    if (detailedEls.length == 0) {
        chkDetailedNodes.attr("disabled", "disabled");
    }
    else if (!showDetailedNotes) {
        setDetailedStyle(detailedEls, "none");
    }
}

function chkDetailedNodes_click() {
    var chkDetailedNodes = $("#chkDetailedNotes");
    var checked = chkDetailedNodes.is(":checked");
    var display = checked ? "block" : "none";

    var detailedEls = $(".detailed");
    setDetailedStyle(detailedEls, display);

    window.external.chkDetailedNodes_Changed(checked);
}

function setDetailedStyle(elements, display) {
    for (var i = 0; i < elements.length; i++) {
        var el = elements[i];
        setDetailedParentStyle(el, display);
    }
}

function setDetailedParentStyle(el, display) {
    var selector = "." + el.className.replace(" ", ".")
    if ($(el.parentNode).children(selector).length == 1)
        setDetailedParentStyle(el.parentNode, display);
    else
        el.style.display = display;
}
