// ============================================================
// EXCEL GRADIENT TOOL
// ============================================================

let selectedFiles = [];


// ============================================================
// ELEMENTS
// ============================================================

const excelFile = document.getElementById("excelFile");
const fileList = document.getElementById("fileList");

const sheetSelection =
    document.getElementById("sheetSelection");
const cellRange = document.getElementById("cellRange");
const lowValue = document.getElementById("lowValue");
const highValue = document.getElementById("highValue");

const lowColor = document.getElementById("lowColor");
const highColor = document.getElementById("highColor");

const applyGradient = document.getElementById("applyGradient");
const addRange = document.getElementById("addRange");
const extraRulesContainer =
    document.getElementById("extraRulesContainer");

let ruleCount = 1;


// ============================================================
// HELPER — COLOR
// ============================================================

function toARGB(hex) {

    const clean = hex.replace("#", "").toUpperCase();

    return "FF" + clean;
}


// ============================================================
// DISPLAY SELECTED FILES
// ============================================================

function renderFileList() {

    fileList.innerHTML = "";

    if (selectedFiles.length === 0) {

        fileList.innerHTML = `
            <div class="empty-file-message">
                No files selected
            </div>
        `;

        return;
    }


    selectedFiles.forEach(function (file, index) {

        const fileItem = document.createElement("div");

        fileItem.className = "file-item";


        fileItem.innerHTML = `

            <span class="file-icon">
                📄
            </span>

            <span
                class="file-item-name"
                title="${file.name}"
            >
                ${file.name}
            </span>

            <button
                type="button"
                class="remove-file"
                data-index="${index}"
                title="Remove file"
            >
                ×
            </button>

        `;


        fileList.appendChild(fileItem);

    });
}


// ============================================================
// SELECT FILES
// ============================================================

excelFile.addEventListener("change", function () {

    const newFiles = Array.from(excelFile.files);


    newFiles.forEach(function (file) {

        const alreadyExists =
            selectedFiles.some(function (existingFile) {

                return (
                    existingFile.name === file.name &&
                    existingFile.size === file.size &&
                    existingFile.lastModified === file.lastModified
                );

            });


        if (!alreadyExists) {

            selectedFiles.push(file);

        }

    });


    renderFileList();


    // Reset browser file input.
    // This lets the user choose the same file again later
    // after removing it.

    excelFile.value = "";

});


// ============================================================
// REMOVE FILE
// ============================================================

fileList.addEventListener("click", function (event) {

    const removeButton =
        event.target.closest(".remove-file");


    if (!removeButton) {

        return;

    }


    const index =
        Number(removeButton.dataset.index);


    selectedFiles.splice(index, 1);


    renderFileList();

});


// ============================================================
// ADD ANOTHER GRADIENT RULE
// ============================================================

addRange.addEventListener("click", function () {

    ruleCount++;


    const newRule =
        document.createElement("div");


    newRule.className = "gradient-rule";


    newRule.innerHTML = `

        <div class="rule-number">
            ${ruleCount}
        </div>


        <div class="field">

            <label>
                Cell range
            </label>

            <input
                type="text"
                placeholder="Example: D1:D30"
            >

        </div>


        <div class="field">

            <label>
                Low threshold
            </label>

            <input
                type="number"
                value="6"
            >

        </div>


        <div class="color-area">

            <div class="color-labels">

                <span>
                    Low
                </span>

                <span>
                    High
                </span>

            </div>


            <div class="gradient-preview">

                <input
                    type="color"
                    value="#5B8CCB"
                    title="Choose low color"
                >

                <div class="gradient-line"></div>

                <input
                    type="color"
                    value="#E05252"
                    title="Choose high color"
                >

            </div>

        </div>


        <div class="field">

            <label>
                High threshold
            </label>

            <input
                type="number"
                value="20"
            >

        </div>

    `;


    extraRulesContainer.appendChild(newRule);

});


// ============================================================
// COLLECT RULES FROM WEBSITE
// ============================================================

function collectRules() {

    const rules = [];


    // FIRST RULE

    rules.push({

        range: cellRange.value.trim(),

        low: Number(lowValue.value),

        high: Number(highValue.value),

        lowColor: lowColor.value,

        highColor: highColor.value

    });


    // EXTRA RULES

    const extraRules =
        document.querySelectorAll(".gradient-rule");


    extraRules.forEach(function (rule) {

        const inputs =
            rule.querySelectorAll("input");


        rules.push({

            range: inputs[0].value.trim(),

            low: Number(inputs[1].value),

            high: Number(inputs[4].value),

            lowColor: inputs[2].value,

            highColor: inputs[3].value

        });

    });


    return rules;

}


// ============================================================
// VALIDATE RULE
// ============================================================

function ruleIsValid(rule) {

    if (!rule.range) {

        return false;

    }


    if (!rule.range.includes(":")) {

        return false;

    }


    if (!Number.isFinite(rule.low)) {

        return false;

    }


    if (!Number.isFinite(rule.high)) {

        return false;

    }


    if (rule.low >= rule.high) {

        return false;

    }


    return true;

}


// ============================================================
// FIND NUMERIC VALUES IN RANGE
// ============================================================

function getRangeInformation(worksheet, rangeText) {

    const parts =
        rangeText.split(":");


    if (parts.length !== 2) {

        throw new Error(
            "Invalid cell range: " + rangeText
        );

    }


    const startCell =
        worksheet.getCell(parts[0]);


    const endCell =
        worksheet.getCell(parts[1]);


    const values = [];


    for (
        let row = startCell.row;
        row <= endCell.row;
        row++
    ) {

        for (
            let column = startCell.col;
            column <= endCell.col;
            column++
        ) {

            const cell =
                worksheet.getCell(row, column);


            if (typeof cell.value === "number") {

                values.push(cell.value);

            }

        }

    }


    return {

        startCell,
        endCell,
        values

    };

}


// ============================================================
// APPLY CONDITIONAL FORMATTING
// ============================================================

function applyRuleToWorksheet(worksheet, rule) {

    const info =
        getRangeInformation(
            worksheet,
            rule.range
        );

    if (info.values.length === 0) {

        console.warn(
            "No numeric values found in",
            rule.range,
            "on",
            worksheet.name
        );

        return;
    }

    const minValue =
        Math.min(...info.values);

    const maxValue =
        Math.max(...info.values);

    const lowARGB =
        toARGB(rule.lowColor);

    const highARGB =
        toARGB(rule.highColor);


    // ------------------------------------------
    // Build TWO NON-OVERLAPPING cell lists
    // ------------------------------------------

    const lowCells = [];
    const highCells = [];


    for (
        let row = info.startCell.row;
        row <= info.endCell.row;
        row++
    ) {

        for (
            let column = info.startCell.col;
            column <= info.endCell.col;
            column++
        ) {

            const cell =
                worksheet.getCell(
                    row,
                    column
                );

            const value =
                cell.value;


            if (
                typeof value !== "number"
            ) {
                continue;
            }


            if (value <= rule.low) {

                lowCells.push(
                    cell.address
                );

            } else if (
                value >= rule.high
            ) {

                highCells.push(
                    cell.address
                );
            }
        }
    }


    console.log(
        "LOW CELLS:",
        lowCells
    );

    console.log(
        "HIGH CELLS:",
        highCells
    );


    // ------------------------------------------
    // LOW GRADIENT
    // ------------------------------------------

    if (lowCells.length > 0) {

        worksheet.addConditionalFormatting({

            ref:
                lowCells.join(" "),

            rules: [
                {
                    type:
                        "colorScale",

                    cfvo: [
                        {
                            type:
                                "num",
                            value:
                                minValue
                        },
                        {
                            type:
                                "num",
                            value:
                                rule.low
                        }
                    ],

                    color: [
                        {
                            argb:
                                lowARGB
                        },
                        {
                            argb:
                                "FFF7FBFF"
                        }
                    ]
                }
            ]
        });
    }


    // ------------------------------------------
    // HIGH GRADIENT
    // ------------------------------------------

    if (highCells.length > 0) {

        worksheet.addConditionalFormatting({

            ref:
                highCells.join(" "),

            rules: [
                {
                    type:
                        "colorScale",

                    cfvo: [
                        {
                            type:
                                "num",
                            value:
                                rule.high
                        },
                        {
                            type:
                                "num",
                            value:
                                maxValue
                        }
                    ],

                    color: [
                        {
                            argb:
                                "FFFFF7F7"
                        },
                        {
                            argb:
                                highARGB
                        }
                    ]
                }
            ]
        });
    }
}


// ============================================================
// DOWNLOAD WORKBOOK
// ============================================================

async function downloadWorkbook(workbook, originalName) {

    const outputBuffer =
        await workbook.xlsx.writeBuffer();


    const blob =
        new Blob(

            [outputBuffer],

            {
                type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }

        );


    const downloadUrl =
        URL.createObjectURL(blob);


    const downloadLink =
        document.createElement("a");


    downloadLink.href =
        downloadUrl;


    const newFileName =
        originalName.replace(
            /\.xlsx$/i,
            "_gradient.xlsx"
        );


    downloadLink.download =
        newFileName;


    document.body.appendChild(
        downloadLink
    );


    downloadLink.click();


    downloadLink.remove();


    setTimeout(function () {

        URL.revokeObjectURL(
            downloadUrl
        );

    }, 1000);

}
function parseSheetSelection(text, totalSheets) {

    const cleaned = text.trim();

    // Blank = every sheet
    if (cleaned === "") {

        return Array.from(
            { length: totalSheets },
            (_, index) => index + 1
        );
    }

    const selectedSheets = new Set();

    const parts = cleaned.split(",");

    parts.forEach(function (part) {

        part = part.trim();

        if (!part) {
            return;
        }

        // RANGE, for example 6:9
        if (part.includes(":")) {

            const [startText, endText] =
                part.split(":");

            const start = Number(startText.trim());
            const end = Number(endText.trim());

            if (
                !Number.isInteger(start) ||
                !Number.isInteger(end)
            ) {
                return;
            }

            const first = Math.min(start, end);
            const last = Math.max(start, end);

            for (
                let number = first;
                number <= last;
                number++
            ) {

                if (
                    number >= 1 &&
                    number <= totalSheets
                ) {
                    selectedSheets.add(number);
                }
            }

        }

        // SINGLE SHEET, for example 4
        else {

            const number = Number(part);

            if (
                Number.isInteger(number) &&
                number >= 1 &&
                number <= totalSheets
            ) {
                selectedSheets.add(number);
            }
        }
    });

    return Array.from(selectedSheets)
        .sort((a, b) => a - b);
}

// ============================================================
// PROCESS ONE EXCEL FILE
// ============================================================

async function processExcelFile(file, rules) {

    console.log(
        "Processing:",
        file.name
    );


    // NEW workbook for EVERY uploaded file.

    const workbook =
        new ExcelJS.Workbook();


    const buffer =
        await file.arrayBuffer();


    await workbook.xlsx.load(buffer);


    // Apply SAME website rules to every worksheet.

    const selectedSheetNumbers =
    parseSheetSelection(
        sheetSelection.value,
        workbook.worksheets.length
    );

console.log(
    "Applying gradient to sheets:",
    selectedSheetNumbers
);

workbook.worksheets.forEach(
    function (worksheet, index) {

        const sheetNumber = index + 1;

        if (
            !selectedSheetNumbers.includes(
                sheetNumber
            )
        ) {
            return;
        }

        rules.forEach(function (rule) {

            applyRuleToWorksheet(
                worksheet,
                rule
            );

        });
    }
);


    await downloadWorkbook(
        workbook,
        file.name
    );

}


// ============================================================
// APPLY GRADIENT BUTTON
// ============================================================

applyGradient.addEventListener(
    "click",
    async function () {


        // ----------------------------------------------------
        // CHECK FILES
        // ----------------------------------------------------

        if (selectedFiles.length === 0) {

            alert(
                "Please select at least one Excel file."
            );

            return;

        }


        // ----------------------------------------------------
        // GET RULES
        // ----------------------------------------------------

        const rules =
            collectRules();


        const invalidRule =
            rules.find(function (rule) {

                return !ruleIsValid(rule);

            });


        if (invalidRule) {

            alert(
                "Please check your gradient rules. " +
                "Every rule needs a valid range, " +
                "low threshold, and high threshold."
            );

            return;

        }


        console.log(
            "Gradient rules:",
            rules
        );


        // ----------------------------------------------------
        // DISABLE BUTTON WHILE WORKING
        // ----------------------------------------------------

        const oldButtonText =
            applyGradient.textContent;


        applyGradient.disabled =
            true;


        applyGradient.textContent =
            "Processing...";


        try {


            // ------------------------------------------------
            // PROCESS EVERY FILE SEPARATELY
            // ------------------------------------------------

            for (
                const file of selectedFiles
            ) {

                await processExcelFile(
                    file,
                    rules
                );


                /*
                    Small delay between browser downloads.

                    This helps when many files are selected.
                */

                await new Promise(function (resolve) {

                    setTimeout(
                        resolve,
                        300
                    );

                });

            }


            console.log(
                "Finished processing",
                selectedFiles.length,
                "file(s)."
            );


        }

        catch (error) {

            console.error(
                "Gradient processing failed:",
                error
            );


            alert(
                "Something went wrong while processing the Excel file. " +
                "Open F12 → Console and send me the red error."
            );

        }

        finally {


            applyGradient.disabled =
                false;


            applyGradient.textContent =
                oldButtonText;

        }

    }
);