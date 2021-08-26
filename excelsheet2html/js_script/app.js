const tbl = document.querySelector(".[PREF]table");

function _get_formatNumbers(format, precission) {
  if (format === "None") {
    return (x) => x.toFixed(parseInt(precission, 10));
  }
  if (format === "comma") {
    return (x) =>
      x
        .toFixed(parseInt(precission, 10))
        .toString()
        .replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }
  if (format === "space") {
    return (x) =>
      x
        .toFixed(parseInt(precission, 10))
        .toString()
        .replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  }

  return (x) => x;
}

const format_number = _get_formatNumbers("[FORMAT]", "[PRECISION]");

tbl.addEventListener("dblclick", (e) => {
  const el = e.target;
  if (el.classList.contains("[PREF]values")) {
    el.contentEditable = true;
    setTimeout(function () {
      el.focus();
      if (el.dataset["isFormula"]) {
        el.innerText = table_data.secondary_cells[el.id];
        const f = table_data.secondary_cells_1.filter((x) => x.parent === el.id);
        for (const k of f) {
          document.getElementById(k.child).classList.add("[PREF]formula_childs");
        }
      }

      if (!el.dataset["isFormula"]) {
        el.innerText = table_data.all_cells[el.id];
      }
    }, 0);
  }
});

tbl.addEventListener("focusout", (e) => {
  const el = e.target;
  if (el.classList.contains("[PREF]values")) {
    el.contentEditable = false;
    if (el.dataset["isFormula"]) {
      el.innerText = format_number(table_data.all_cells[el.id]);
      const f = table_data.secondary_cells_1.filter((x) => x.parent === el.id);
      for (const k of f) {
        document.getElementById(k.child).classList.remove("[PREF]formula_childs");
      }
    }

    if (!el.dataset["isFormula"]) {
      el.innerText = format_number(table_data.all_cells[el.id]);
    }
  }
});

function recalculate_values() {
  const new_values = {};
  for (const k in table_data.primary_cells) {
    new_values[k] = table_data.all_cells[k];
  }
  let iter = 1;
  while (true) {
    let f = table_data.secondary_cells_2.filter((el) => el.iteration === iter);
    if (f.length === 0) {
      break;
    }
    for (const e of f) {
      new_values[e.parent] = (new_values[e.parent] || 0) + new_values[e.child] * e.multiplier;
    }
    iter++;
  }
  for (const k in new_values) {
    if (!k.startsWith("XXX")) {
      document.getElementById(k).innerText = format_number(new_values[k]);
    }
  }
  table_data.all_cells = new_values;
}

tbl.addEventListener("keypress", (e) => {
  const el = e.target;
  const before_modification = table_data.all_cells[el.id];
  if (el.classList.contains("[PREF]values")) {
    if (e.key === "Enter") {
      if (!el.dataset["isFormula"]) {
        table_data.all_cells[el.id] = eval(el.innerText.replace("=", "+")) * 1;
        const after_modyfication = eval(el.innerText.replace("=", "+")) * 1;
        // console.log("Przed: ", before_modification, "   Po:", after_modyfication);
        if (before_modification !== after_modyfication) {
          recalculate_values();
        }
        //el.innerText = table_data.formula_cells[el.id];
      }
      el.contentEditable = false;
    }
  }
});

function getStyleVal(elm, css) {
  return window.getComputedStyle(elm, null).getPropertyValue(css);
}

function paddingDiff(col) {
  if (getStyleVal(col, "box-sizing") == "border-box") {
    return 0;
  }

  var padLeft = getStyleVal(col, "padding-left");
  var padRight = getStyleVal(col, "padding-right");
  return parseInt(padLeft) + parseInt(padRight);
}

function setListeners(div) {
  var pageX, curCol, nxtCol, curColWidth, nxtColWidth;
  div.addEventListener("mousedown", function (e) {
    curCol = e.target.parentElement;
    pageX = e.pageX;
    curColLetter = curCol.innerText;
    selected_class = ".[PREF]col_" + curColLetter;

    var padding = paddingDiff(curCol);

    curColWidth = curCol.offsetWidth - padding;
    cur_table_width = document.querySelector(".[PREF]table").offsetWidth;
    if (nxtCol) nxtColWidth = nxtCol.offsetWidth - padding;
  });

  div.addEventListener("mouseover", function (e) {
    e.target.style.borderRight = "2px solid #0000ff";
  });

  div.addEventListener("mouseout", function (e) {
    e.target.style.borderRight = "";
  });

  document.addEventListener("mousemove", function (e) {
    if (curCol) {
      var diffX = e.pageX - pageX;
      //console.log((curColWidth + diffX))
      document.querySelector(selected_class).style.width = curColWidth + diffX + "px";
      document.querySelector(".[PREF]table").style.width = cur_table_width + diffX + "px";
    }
  });

  document.addEventListener("mouseup", function (e) {
    curCol = undefined;
    nxtCol = undefined;
    pageX = undefined;
    nxtColWidth = undefined;
    curColWidth = undefined;
  });
}

const all_resize = document.querySelectorAll(".[PREF]table th div");
const arr = Array.from(all_resize);
for (let i = 1; i < arr.length; i++) {
  setListeners(arr[i]);
}

table_data["all_cells"] = {
  ...table_data.primary_cells,
  ...table_data.formula_cells,
};
table_data["all_cells_ory"] = {
  ...table_data.primary_cells,
  ...table_data.formula_cells,
};
table_data["changed_cells"] = {};

for (const k in table_data.all_cells_ory) {
  if (!k.startsWith("XXX")) {
    document.getElementById(k).innerText = format_number(table_data.all_cells_ory[k]);
  }
}
