// Pointer-driven pinch/pan zoom for content this page renders itself.
// Cross-origin iframes (Google Drive, SharePoint) are deliberately excluded:
// their touch events never reach this document, so Google's own viewer owns
// gestures there and cannot be overridden from here.
(function () {
    "use strict";

    const MIN_SCALE = 1;
    const MAX_SCALE = 6;
    const DOUBLE_TAP_MS = 320;

    function attach(viewport, content, onScaleChange) {
        const state = { scale: 1, x: 0, y: 0 };
        const pointers = new Map();
        let pinchDistance = 0;
        let pinchScale = 1;
        let lastTapAt = 0;

        content.style.transformOrigin = "0 0";
        viewport.style.touchAction = "none";
        viewport.style.overflow = "hidden";

        function clamp() {
            const box = viewport.getBoundingClientRect();
            const width = content.offsetWidth * state.scale;
            const height = content.offsetHeight * state.scale;
            // Centre content that fits; otherwise keep the panned edges inside the viewport.
            state.x = width <= box.width ? (box.width - width) / 2 : Math.min(0, Math.max(box.width - width, state.x));
            state.y = height <= box.height ? (box.height - height) / 2 : Math.min(0, Math.max(box.height - height, state.y));
        }

        function apply() {
            clamp();
            content.style.transform = `translate(${state.x}px, ${state.y}px) scale(${state.scale})`;
            if (onScaleChange) onScaleChange(state.scale);
        }

        function zoomTo(nextScale, clientX, clientY) {
            const box = viewport.getBoundingClientRect();
            const anchorX = clientX - box.left;
            const anchorY = clientY - box.top;
            const bounded = Math.min(MAX_SCALE, Math.max(MIN_SCALE, nextScale));
            const ratio = bounded / state.scale;
            // Keep the point under the fingers (or cursor) fixed while scaling.
            state.x = anchorX - (anchorX - state.x) * ratio;
            state.y = anchorY - (anchorY - state.y) * ratio;
            state.scale = bounded;
            apply();
        }

        function centreOfPointers() {
            const points = [...pointers.values()];
            return {
                x: points.reduce((total, point) => total + point.x, 0) / points.length,
                y: points.reduce((total, point) => total + point.y, 0) / points.length
            };
        }

        function distanceOfPointers() {
            const [first, second] = [...pointers.values()];
            return Math.hypot(first.x - second.x, first.y - second.y);
        }

        viewport.addEventListener("pointerdown", (event) => {
            viewport.setPointerCapture(event.pointerId);
            pointers.set(event.pointerId, { x: event.clientX, y: event.clientY });
            if (pointers.size === 2) {
                pinchDistance = distanceOfPointers();
                pinchScale = state.scale;
            }
        });

        viewport.addEventListener("pointermove", (event) => {
            const previous = pointers.get(event.pointerId);
            if (!previous) return;
            const current = { x: event.clientX, y: event.clientY };
            pointers.set(event.pointerId, current);

            if (pointers.size >= 2) {
                const distance = distanceOfPointers();
                if (pinchDistance > 0) {
                    const centre = centreOfPointers();
                    zoomTo(pinchScale * (distance / pinchDistance), centre.x, centre.y);
                }
                return;
            }

            state.x += current.x - previous.x;
            state.y += current.y - previous.y;
            apply();
            event.preventDefault();
        });

        function releasePointer(event) {
            pointers.delete(event.pointerId);
            if (pointers.size < 2) pinchDistance = 0;
        }

        viewport.addEventListener("pointerup", (event) => {
            releasePointer(event);
            const now = Date.now();
            if (now - lastTapAt < DOUBLE_TAP_MS) {
                zoomTo(state.scale > 1 ? 1 : 2.5, event.clientX, event.clientY);
                lastTapAt = 0;
                return;
            }
            lastTapAt = now;
        });
        viewport.addEventListener("pointercancel", releasePointer);
        viewport.addEventListener("pointerleave", releasePointer);

        viewport.addEventListener("wheel", (event) => {
            if (event.ctrlKey) {
                // Trackpad pinch and ctrl+wheel arrive here as a wheel event.
                zoomTo(state.scale * (event.deltaY < 0 ? 1.12 : 1 / 1.12), event.clientX, event.clientY);
            } else {
                state.x -= event.deltaX;
                state.y -= event.deltaY;
                apply();
            }
            event.preventDefault();
        }, { passive: false });

        apply();

        return {
            get scale() { return state.scale; },
            zoomBy(factor) {
                const box = viewport.getBoundingClientRect();
                zoomTo(state.scale * factor, box.left + box.width / 2, box.top + box.height / 2);
            },
            reset() {
                state.scale = 1;
                state.x = 0;
                state.y = 0;
                apply();
            },
            refresh: apply
        };
    }

    window.RESOURCEFIT_ZOOM = { attach, MIN_SCALE, MAX_SCALE };
}());
