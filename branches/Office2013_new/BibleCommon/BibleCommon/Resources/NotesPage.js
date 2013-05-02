$(function () {
    initDetailedNotes();
});

function initDetailedNotes() {
    var chkDetailedNodes = $("#chkDetailedNotes");
    var detailedEls = $(".detailed");
    if (detailedEls.count == 0) {
        chkDetailedNodes.attr("disabled", "disabled");
    }
}