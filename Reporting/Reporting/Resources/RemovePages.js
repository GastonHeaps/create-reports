function RemovePagesWithoutSection(stringToSearchFor) {

    var pageArray = [];
    var isSearchingForTotals = false;

    stringToSearchFor = stringToSearchFor.toUpperCase().split(" ");

    for (var p = 0; p < this.numPages; p++) {

        var pageContainsSection = isSearchingForTotals;

        // iterate through all words on the page
        for (var n = 0; n < this.getPageNumWords(p) ; n++) {

            if (isSearchingForTotals) {
                if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                    isSearchingForTotals = false;
                }
            } else {
                var strContainsWords;

                for (var w = 0; w < stringToSearchFor.length; w++) {
                    if (stringToSearchFor[w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                        strContainsWords = true;
                    } else {
                        strContainsWords = false;
                        break;
                    }
                }
                if (strContainsWords) {
                    isSearchingForTotals = true;
                    pageContainsSection = true;
                }
            }
        }

        if (!pageContainsSection) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}
function RemovePagesWithoutMTDSection(stringToSearchFor, dateToSearchFor) {

    var pageArray = [];
    var pagesThatContainSection;
    var isSearchingForDate = false;
    var isSearchingForTotals = false;

    stringToSearchFor = stringToSearchFor.toUpperCase().split(" ");

    for (var p = 0; p < this.numPages; p++) {

        var sectionsFoundOnPage;
        var pageContainsSection = isSearchingForTotals;

        if (isSearchingForTotals) {
            sectionsFoundOnPage = 1;
            pagesThatContainSection++;
        } else if (!isSearchingForTotals) {
            sectionsFoundOnPage = 0;
            pagesThatContainSection = 0;
        }

        // iterate through all words on the page
        for (var n = 0; n < this.getPageNumWords(p) ; n++) {

            if (isSearchingForTotals) {
                var firstThreeChar = this.getPageNthWord(p, n, false).substring(0, 3).toUpperCase();

                if (firstThreeChar == "MON" || firstThreeChar == "TUE" || firstThreeChar == "WED" || firstThreeChar == "THU" || firstThreeChar == "FRI" || firstThreeChar == "SAT" || firstThreeChar == "SUN") {
                    if (parseInt(this.getPageNthWord(p, n + 1, false)) >= dateToSearchFor) {
                        isSearchingForDate = false;
                    }
                } else if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                    if (isSearchingForDate) {
                        if (pagesThatContainSection > 1) {
                            // Delete false-positive section
                            pageContainsSection = false;
                            for (var i = 1; i < pagesThatContainSection; i++) {
                                pageArray.push(p - i);
                            }
                        } else if (pagesThatContainSection == 1) {
                            if (sectionsFoundOnPage == 1) {
                                pageContainsSection = false;
                            }
                        }
                        sectionsFoundOnPage--;
                        isSearchingForDate = false;
                        isSearchingForTotals = false;
                    } else if (!isSearchingForDate) {
                        isSearchingForTotals = false;
                    }
                    pagesThatContainSection = 0;
                }
            } else {
                var strContainsWords;

                for (var w = 0; w < stringToSearchFor.length; w++) {
                    if (stringToSearchFor[w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                        strContainsWords = true;
                    } else {
                        strContainsWords = false;
                        break;
                    }
                }
                if (strContainsWords) {
                    pagesThatContainSection = 1;
                    sectionsFoundOnPage++;
                    pageContainsSection = true;
                    isSearchingForDate = true;
                    isSearchingForTotals = true;
                }
            }
        }

        if (!pageContainsSection || (isSearchingForDate && pagesThatContainSection > 1)) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}

function RemovePagesWithoutSectionWithExactText(stringToSearchFor) {

    var pageArray = [];
    var isSearchingForTotals = false;

    stringToSearchFor = stringToSearchFor.toUpperCase().split(" ");

    for (var p = 0; p < this.numPages; p++) {

        var pageContainsSection = isSearchingForTotals;

        // iterate through all words on the page
        LoopWords:
            for (var n = 0; n < this.getPageNumWords(p) ; n++) {

                if (isSearchingForTotals) {
                    if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                        isSearchingForTotals = false;
                    }
                } else {
                    var strContainsWords;

                    for (var w = 0; w < stringToSearchFor.length; w++) {
                        if (stringToSearchFor[w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                            strContainsWords = true;
                        } else {
                            strContainsWords = false;
                            break;
                        }
                    }
                    if (strContainsWords) {
                        var firstWordQuads = this.getPageNthWordQuads(p, n);
                        var preFirstWordQuads = this.getPageNthWordQuads(p, n - 1);
                        var postLastWordQuads = this.getPageNthWordQuads(p, n + stringToSearchFor.length);

                        if (preFirstWordQuads[0][1] == firstWordQuads[0][1]) {
                            continue LoopWords;
                        }
                        if (postLastWordQuads[0][1] == firstWordQuads[0][1]) {
                            continue LoopWords;
                        }

                        isSearchingForTotals = true;
                        pageContainsSection = true;
                    }
                }
            }

        if (!pageContainsSection) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}
function RemovePagesWithoutMTDSectionWithExactText(stringToSearchFor, dateToSearchFor) {

    var pageArray = [];
    var pagesThatContainSection;
    var isSearchingForDate = false;
    var isSearchingForTotals = false;

    stringToSearchFor = stringToSearchFor.toUpperCase().split(" ");

    LoopPages:
        for (var p = 0; p < this.numPages; p++) {

            var sectionsFoundOnPage;
            var pageContainsSection = isSearchingForTotals;

            if (isSearchingForTotals) {
                sectionsFoundOnPage = 1;
                pagesThatContainSection++;
            } else if (!isSearchingForTotals) {
                sectionsFoundOnPage = 0;
                pagesThatContainSection = 0;
            }

            // iterate through all words on the page
            LoopWords:
                for (var n = 0; n < this.getPageNumWords(p) ; n++) {

                    if (isSearchingForTotals) {
                        var firstThreeChar = this.getPageNthWord(p, n, false).substring(0, 3).toUpperCase();

                        if (firstThreeChar == "MON" || firstThreeChar == "TUE" || firstThreeChar == "WED" || firstThreeChar == "THU" || firstThreeChar == "FRI" || firstThreeChar == "SAT" || firstThreeChar == "SUN") {
                            if (parseInt(this.getPageNthWord(p, n + 1, false)) >= dateToSearchFor) {
                                isSearchingForDate = false;
                            }
                        } else if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                            if (isSearchingForDate) {
                                if (pagesThatContainSection > 1) {
                                    // Delete false-positive section
                                    pageContainsSection = false;
                                    for (var i = 1; i < pagesThatContainSection; i++) {
                                        pageArray.push(p - i);
                                    }
                                } else if (pagesThatContainSection == 1) {
                                    if (sectionsFoundOnPage == 1) {
                                        pageContainsSection = false;
                                    }
                                }
                                sectionsFoundOnPage--;
                                isSearchingForDate = false;
                                isSearchingForTotals = false;
                            } else if (!isSearchingForDate) {
                                isSearchingForTotals = false;
                            }
                            pagesThatContainSection = 0;
                        }
                    } else {
                        var strContainsWords;

                        for (var w = 0; w < stringToSearchFor.length; w++) {
                            if (stringToSearchFor[w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                                strContainsWords = true;
                            } else {
                                strContainsWords = false;
                                break;
                            }
                        }
                        if (strContainsWords) {
                            var firstWordQuads = this.getPageNthWordQuads(p, n);
                            var preFirstWordQuads = this.getPageNthWordQuads(p, n - 1);
                            var postLastWordQuads = this.getPageNthWordQuads(p, n + stringToSearchFor.length);

                            if (preFirstWordQuads[0][1] == firstWordQuads[0][1]) {
                                continue LoopWords;
                            }
                            if (postLastWordQuads[0][1] == firstWordQuads[0][1]) {
                                continue LoopWords;
                            }

                            pagesThatContainSection = 1;
                            sectionsFoundOnPage++;
                            pageContainsSection = true;
                            isSearchingForDate = true;
                            isSearchingForTotals = true;
                        }
                    }
                }

            if (!pageContainsSection || (isSearchingForDate && pagesThatContainSection > 1)) {
                pageArray.push(p);
            }
        }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}

function RemovePagesWithoutSections(stringsToSearchFor) {

    var pageArray = [];
    var isSearchingForTotals = false;

    for (var i = 0; i < stringsToSearchFor.length; i++) {
        stringsToSearchFor[i] = stringsToSearchFor[i].toUpperCase().split(" ");
    }

    for (var p = 0; p < this.numPages; p++) {

        var pageContainsSection = isSearchingForTotals;

        // iterate through all words on the page
        for (var n = 0; n < this.getPageNumWords(p) ; n++) {

            if (isSearchingForTotals) {
                if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                    isSearchingForTotals = false;
                }
            } else {
                // iterate through strings
                LoopStrings:
                    for (var s = 0; s < stringsToSearchFor.length; s++) {
                        var strContainsWords;

                        // iterate through words in string
                        for (var w = 0; w < stringsToSearchFor[s].length; w++) {
                            if (stringsToSearchFor[s][w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                                strContainsWords = true;
                            } else {
                                strContainsWords = false;
                                continue LoopStrings;
                            }
                        }
                        if (strContainsWords) {
                            pageContainsSection = true;
                            isSearchingForTotals = true;
                            break;
                        }
                    }
            }

        }
        if (!pageContainsSection) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}
function RemovePagesWithoutMTDSections(stringsToSearchFor, dateToSearchFor) {

    var pageArray = [];
    var pagesThatContainSection;
    var isSearchingForDate = false;
    var isSearchingForTotals = false;

    for (var i = 0; i < stringsToSearchFor.length; i++) {
        stringsToSearchFor[i] = stringsToSearchFor[i].toUpperCase().split(" ");
    }

    for (var p = 0; p < this.numPages; p++) {

        var sectionsFoundOnPage;
        var pageContainsSection = isSearchingForTotals;

        if (isSearchingForTotals) {
            sectionsFoundOnPage = 1;
            pagesThatContainSection++;
        } else if (!isSearchingForTotals) {
            sectionsFoundOnPage = 0;
            pagesThatContainSection = 0;
        }

        // iterate through all words on the page
        for (var n = 0; n < this.getPageNumWords(p) ; n++) {

            if (isSearchingForTotals) {
                var firstThreeChar = this.getPageNthWord(p, n, false).substring(0, 3).toUpperCase();

                if (firstThreeChar == "MON" || firstThreeChar == "TUE" || firstThreeChar == "WED" || firstThreeChar == "THU" || firstThreeChar == "FRI" || firstThreeChar == "SAT" || firstThreeChar == "SUN") {
                    if (parseInt(this.getPageNthWord(p, n + 1, false)) >= dateToSearchFor) {
                        isSearchingForDate = false;
                    }
                } else if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                    if (isSearchingForDate) {
                        if (pagesThatContainSection > 1) {
                            // Delete false-positive section
                            pageContainsSection = false;
                            for (var i = 1; i < pagesThatContainSection; i++) {
                                pageArray.push(p - i);
                            }
                        } else if (pagesThatContainSection == 1) {
                            if (sectionsFoundOnPage == 1) {
                                pageContainsSection = false;
                            }
                        }
                        sectionsFoundOnPage--;
                        isSearchingForDate = false;
                        isSearchingForTotals = false;
                    } else if (!isSearchingForDate) {
                        isSearchingForTotals = false;
                    }
                    pagesThatContainSection = 0;
                }
            } else {
                // iterate through strings
                LoopStrings:
                    for (var s = 0; s < stringsToSearchFor.length; s++) {
                        var strContainsWords;

                        // iterate through words in string
                        for (var w = 0; w < stringsToSearchFor[s].length; w++) {
                            if (stringsToSearchFor[s][w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                                strContainsWords = true;
                            } else {
                                strContainsWords = false;
                                continue LoopStrings;
                            }
                        }
                        if (strContainsWords) {
                            pagesThatContainSection = 1;
                            sectionsFoundOnPage++;
                            pageContainsSection = true;
                            isSearchingForDate = true;
                            isSearchingForTotals = true;
                            break;
                        }
                    }
            }

        }
        if (!pageContainsSection || (isSearchingForDate && pagesThatContainSection > 1)) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}

function RemovePagesWithoutSectionsWithExactText(stringsToSearchFor) {

    var pageArray = [];
    var isSearchingForTotals = false;

    for (var i = 0; i < stringsToSearchFor.length; i++) {
        stringsToSearchFor[i] = stringsToSearchFor[i].toUpperCase().split(" ");
    }

    for (var p = 0; p < this.numPages; p++) {

        var pageContainsSection = isSearchingForTotals;

        // iterate through all words on the page
        for (var n = 0; n < this.getPageNumWords(p) ; n++) {

            if (isSearchingForTotals) {
                if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                    isSearchingForTotals = false;
                }
            } else {
                // iterate through strings
                LoopStrings:
                    for (var s = 0; s < stringsToSearchFor.length; s++) {
                        var strContainsWords;

                        // iterate through words in string
                        for (var w = 0; w < stringsToSearchFor[s].length; w++) {
                            if (stringsToSearchFor[s][w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                                strContainsWords = true;
                            } else {
                                strContainsWords = false;
                                continue LoopStrings;
                            }
                        }
                        if (strContainsWords) {
                            var firstWordQuads = this.getPageNthWordQuads(p, n);
                            var preFirstWordQuads = this.getPageNthWordQuads(p, n - 1);
                            var postLastWordQuads = this.getPageNthWordQuads(p, n + stringsToSearchFor[s].length);

                            if (preFirstWordQuads[0][1] == firstWordQuads[0][1]) {
                                continue LoopStrings;
                            }
                            if (postLastWordQuads[0][1] == firstWordQuads[0][1]) {
                                continue LoopStrings;
                            }

                            pageContainsSection = true;
                            isSearchingForTotals = true;
                            break;
                        }
                    }
            }

        }
        if (!pageContainsSection) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}
function RemovePagesWithoutMTDSectionsWithExactText(stringsToSearchFor, dateToSearchFor) {

    var pageArray = [];
    var pagesThatContainSection;
    var isSearchingForDate = false;
    var isSearchingForTotals = false;

    for (var i = 0; i < stringsToSearchFor.length; i++) {
        stringsToSearchFor[i] = stringsToSearchFor[i].toUpperCase().split(" ");
    }

    for (var p = 0; p < this.numPages; p++) {

        var sectionsFoundOnPage;
        var pageContainsSection = isSearchingForTotals;

        if (isSearchingForTotals) {
            sectionsFoundOnPage = 1;
            pagesThatContainSection++;
        } else if (!isSearchingForTotals) {
            sectionsFoundOnPage = 0;
            pagesThatContainSection = 0;
        }

        // Iterate through all words on the page
        for (var n = 0; n < this.getPageNumWords(p) ; n++) {

            if (isSearchingForTotals) {
                var firstThreeChar = this.getPageNthWord(p, n, false).substring(0, 3).toUpperCase();

                if (firstThreeChar == "MON" || firstThreeChar == "TUE" || firstThreeChar == "WED" || firstThreeChar == "THU" || firstThreeChar == "FRI" || firstThreeChar == "SAT" || firstThreeChar == "SUN") {
                    if (parseInt(this.getPageNthWord(p, n + 1, false)) >= dateToSearchFor) {
                        isSearchingForDate = false;
                    }
                } else if (this.getPageNthWord(p, n, false).substring(0, 5).toUpperCase() == "TOTAL") {
                    if (isSearchingForDate) {
                        if (pagesThatContainSection > 1) {
                            // Delete false-positive section
                            pageContainsSection = false;
                            for (var i = 1; i < pagesThatContainSection; i++) {
                                pageArray.push(p - i);
                            }
                        } else if (pagesThatContainSection == 1) {
                            if (sectionsFoundOnPage == 1) {
                                pageContainsSection = false;
                            }
                        }
                        sectionsFoundOnPage--;
                        isSearchingForDate = false;
                        isSearchingForTotals = false;
                    } else if (!isSearchingForDate) {
                        isSearchingForTotals = false;
                    }
                    pagesThatContainSection = 0;
                }
            } else {
                // Iterate through strings
                LoopStrings:
                    for (var s = 0; s < stringsToSearchFor.length; s++) {
                        var strContainsWords;

                        // Iterate through words in string
                        for (var w = 0; w < stringsToSearchFor[s].length; w++) {
                            if (stringsToSearchFor[s][w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                                strContainsWords = true;
                            } else {
                                strContainsWords = false;
                                continue LoopStrings;
                            }
                        }
                        if (strContainsWords) {
                            var firstWordQuads = this.getPageNthWordQuads(p, n);
                            var preFirstWordQuads = this.getPageNthWordQuads(p, n - 1);
                            var postLastWordQuads = this.getPageNthWordQuads(p, n + stringsToSearchFor[s].length);

                            if (preFirstWordQuads[0][1] == firstWordQuads[0][1]) {
                                continue LoopStrings;
                            }
                            if (postLastWordQuads[0][1] == firstWordQuads[0][1]) {
                                continue LoopStrings;
                            }

                            pagesThatContainSection = 1;
                            sectionsFoundOnPage++;
                            pageContainsSection = true;
                            isSearchingForDate = true;
                            isSearchingForTotals = true;
                            break;
                        }
                    }
            }

        }
        if (!pageContainsSection || (isSearchingForDate && pagesThatContainSection > 1)) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}

function RemovePagesWithoutText(stringsToSearchFor) {

    var pageArray = [];

    for (var i = 0; i < stringsToSearchFor.length; i++) {
        stringsToSearchFor[i] = stringsToSearchFor[i].toUpperCase().split(" ");
    }

    for (var p = 0; p < this.numPages; p++) {

        var pageContainsText = false;

        // iterate through all words on the page
        LoopPageWords:
            for (var n = 0; n < this.getPageNumWords(p) ; n++) {

                // iterate through strings
                LoopStrings:
                    for (var s = 0; s < stringsToSearchFor.length; s++) {
                        var strContainsWords;

                        // iterate through words in string
                        for (var w = 0; w < stringsToSearchFor[s].length; w++) {
                            if (stringsToSearchFor[s][w] == this.getPageNthWord(p, n + w, false).replace(/^\s\s*/, '').replace(/\s\s*$/, '').toUpperCase()) {
                                strContainsWords = true;
                            } else {
                                strContainsWords = false;
                                continue LoopStrings;
                            }
                        }
                        if (strContainsWords) {
                            pageContainsText = true;
                            break LoopPageWords;
                        }
                    }

            }
        if (!pageContainsText) {
            pageArray.push(p);
        }
    }

    if (pageArray.length > 0 && pageArray.length < this.numPages) {
        for (var n = pageArray.length - 1; n >= 0; n--) {
            this.deletePages({ nStart: pageArray[n] });
        }
        return true;
    } else if (pageArray.length == 0) {
        return true
    } else {
        return false;
    }
}