(function () {
  function blockWithinResults(e) {
    if (e.target.closest && e.target.closest(".results-readonly")) {
      e.preventDefault();
    }
  }
  ["contextmenu", "copy", "cut", "dragstart", "selectstart"].forEach(function (evt) {
    document.addEventListener(evt, blockWithinResults);
  });
})();
