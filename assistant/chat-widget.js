(() => {
  /* ===============================
     BRAND CONFIG
     =============================== */
  const BRAND = {
    name: "Homelyo Assistant",
    primary: "#FF5A5F",
    bubbleIcon: "🤍",
    iframeUrl: "https://www.homelyo.eu/assistant/"
  };

  /* ===============================
     CSS INJECTION
     =============================== */
  const style = document.createElement("style");
  style.innerHTML = `
    #hw-widget {
      position: fixed;
      bottom: 20px;
      right: 20px;
      z-index: 99999;
      font-family: system-ui, -apple-system, sans-serif;
    }

    #hw-bubble {
      width: 56px;
      height: 56px;
      border-radius: 50%;
      border: none;
      background: ${BRAND.primary};
      color: #fff;
      font-size: 24px;
      cursor: pointer;
      box-shadow: 0 8px 24px rgba(0,0,0,.25);
    }

    #hw-window {
      position: absolute;
      bottom: 70px;
      right: 0;
      width: 360px;
      height: 520px;
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 20px 40px rgba(0,0,0,.3);
      overflow: hidden;
      transform: scale(0.9) translateY(10px);
      opacity: 0;
      pointer-events: none;
      transition: all .25s ease;
    }

    #hw-window.open {
      transform: scale(1) translateY(0);
      opacity: 1;
      pointer-events: auto;
    }

    #hw-header {
      height: 48px;
      background: ${BRAND.primary};
      color: #fff;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 12px;
      cursor: move;
      user-select: none;
      font-size: 14px;
    }

    #hw-header button {
      background: none;
      border: none;
      color: #fff;
      font-size: 18px;
      cursor: pointer;
    }

    #hw-window iframe {
      width: 100%;
      height: calc(100% - 48px);
      border: none;
    }

    @media (max-width: 480px) {
      #hw-window {
        width: 90vw;
        height: 70vh;
      }
    }
  `;
  document.head.appendChild(style);

  /* ===============================
     HTML INJECTION
     =============================== */
  const widget = document.createElement("div");
  widget.id = "hw-widget";
  widget.innerHTML = `
    <button id="hw-bubble">${BRAND.bubbleIcon}</button>
    <div id="hw-window">
      <div id="hw-header">
        <span>${BRAND.name}</span>
        <button id="hw-close">✕</button>
      </div>
      <iframe src="${BRAND.iframeUrl}" title="Chat"></iframe>
    </div>
  `;
  document.body.appendChild(widget);

  const bubble = document.getElementById("hw-bubble");
  const windowEl = document.getElementById("hw-window");
  const closeBtn = document.getElementById("hw-close");
  const header = document.getElementById("hw-header");

  /* ===============================
     OPEN / CLOSE
     =============================== */
  bubble.onclick = () => {
    windowEl.classList.add("open");
    bubble.style.display = "none";
    localStorage.setItem("hw-open", "true");
  };

  closeBtn.onclick = () => {
    windowEl.classList.remove("open");
    bubble.style.display = "block";
    localStorage.setItem("hw-open", "false");
  };

  if (localStorage.getItem("hw-open") === "true") {
    bubble.style.display = "none";
    windowEl.classList.add("open");
  }

  /* ===============================
     DRAG / DOCK
     =============================== */
  let isDragging = false;
  let offsetX = 0;
  let offsetY = 0;

  header.addEventListener("mousedown", (e) => {
    isDragging = true;
    const rect = windowEl.getBoundingClientRect();
    offsetX = e.clientX - rect.left;
    offsetY = e.clientY - rect.top;
    windowEl.style.transition = "none";
  });

  document.addEventListener("mousemove", (e) => {
    if (!isDragging) return;
    windowEl.style.left = e.clientX - offsetX + "px";
    windowEl.style.top = e.clientY - offsetY + "px";
    windowEl.style.right = "auto";
    windowEl.style.bottom = "auto";
  });

  document.addEventListener("mouseup", () => {
    isDragging = false;
    windowEl.style.transition = "";
  });
})();
