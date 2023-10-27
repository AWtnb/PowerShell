const compressElem = (query) => {
  Array.from(document.querySelectorAll(query)).forEach((elem) => {
    elem.classList.add("compress")
  });
};
const recoverElem = (query) => {
  Array.from(document.querySelectorAll(query)).forEach((elem) => {
    elem.classList.remove("compress")
  });
};

document.addEventListener(
  "keyup",
  function (e) {
    const pressed = (e.ctrlKey ? "C-" : "") + (e.altKey ? "A-" : "") + (e.shiftKey ? "S-" : "") + e.key.toLowerCase();
    if (pressed == "f") {
      compressElem("ins")
      recoverElem("del")
      return;
    }
    if (pressed == "t") {
      compressElem("del")
      recoverElem("ins")
      return;
    }
    if (pressed == "r") {
      recoverElem("del")
      recoverElem("ins")
      return;
    }
  },
  false
);
