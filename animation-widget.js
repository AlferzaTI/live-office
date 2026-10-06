/* =========================================================
   CUPHEAD ANIMATION WIDGET
========================================================= */

(function () {

    const widget = document.createElement("div");

    widget.id = "cuphead-widget";

    const image = document.createElement("img");

    image.alt = "";

    widget.appendChild(image);

    document.body.appendChild(widget);


    /* =========================================
       FRAMES
    ========================================= */

    const frames = [];

    for (let i = 1; i <= 44; i++) {

        const number = String(i).padStart(4, "0");

        frames.push(
            `img/Cuphead/cuphead_intro_b_${number}.png`
        );

    }


    /* =========================================
       PRECARGAR IMÁGENES
    ========================================= */

    const loadedFrames = [];

    frames.forEach(src => {

        const img = new Image();

        img.src = src;

        loadedFrames.push(img);

    });


    /* =========================================
       ANIMACIÓN
    ========================================= */

    let currentFrame = 0;

    const fps = 14;

    const frameDuration = 1000 / fps;


    function animate() {

        image.src = loadedFrames[currentFrame].src;

        currentFrame++;

        if (currentFrame >= loadedFrames.length) {
            currentFrame = 0;
        }

    }


    animate();

    setInterval(animate, frameDuration);

})();