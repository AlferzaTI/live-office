(function () {

    const widget = document.createElement("div");

    widget.id = "cuphead-widget";

    const image = document.createElement("img");

    image.alt = "Cuphead";

    widget.appendChild(image);

    document.body.appendChild(widget);


    // =========================================
    // CARGAR LOS 44 FRAMES
    // =========================================

    const frames = [];

    for (let i = 1; i <= 44; i++) {

        const number = String(i).padStart(4, "0");

        const frame = new Image();

        frame.src =
            `img/Cuphead/Cuphead/cuphead_intro_b_${number}.png`;

        frames.push(frame);
    }


    // =========================================
    // ANIMACIÓN
    // =========================================

    let currentFrame = 0;

    function animateCuphead() {

        image.src = frames[currentFrame].src;

        currentFrame++;

        if (currentFrame >= frames.length) {
            currentFrame = 0;
        }
    }


    // Primer frame
    animateCuphead();


    // Velocidad
    setInterval(animateCuphead, 80);

})();