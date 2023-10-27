
const setVisibility = (query, display) => {
  Array.from(document.querySelectorAll(query)).forEach((elem) => {
    elem.style.display = display;
  });
};

document.addEventListener(
  "keyup",
  function (e) {
    const pressed = (e.ctrlKey ? "C-" : "") + (e.altKey ? "A-" : "") + (e.shiftKey ? "S-" : "") + e.key.toLowerCase();
    if (pressed == "f") {
      setVisibility("ins", "none")
      setVisibility("del", "inline")
      return;
    }
    if (pressed == "t") {
      setVisibility("del", "none")
      setVisibility("ins", "inline")
      return;
    }
    if (pressed == "r") {
      setVisibility("del", "inline")
      setVisibility("ins", "inline")
      return;
    }
  },
  false
);
