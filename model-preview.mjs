// 3D previews. Mesh formats go through three.js loaders; CAD formats (STEP, IGES,
// BREP) need a geometry kernel, so they are meshed by OpenCascade compiled to wasm.
// Both paths need the bytes, so the host must allow a cross-origin browser download.
import * as THREE from "three";
import { OrbitControls } from "three/addons/controls/OrbitControls.js";

const OCCT_BASE = "https://cdn.jsdelivr.net/npm/occt-import-js@0.0.23/dist/";

const CAD_FORMATS = { stp: "step", step: "step", igs: "iges", iges: "iges", brep: "brep" };
const MESH_LOADERS = {
    stl: ["loaders/STLLoader.js", "STLLoader"],
    obj: ["loaders/OBJLoader.js", "OBJLoader"],
    ply: ["loaders/PLYLoader.js", "PLYLoader"],
    "3mf": ["loaders/3MFLoader.js", "ThreeMFLoader"],
    glb: ["loaders/GLTFLoader.js", "GLTFLoader"],
    gltf: ["loaders/GLTFLoader.js", "GLTFLoader"]
};

export function isModelExtension(extension) {
    return Boolean(CAD_FORMATS[extension] || MESH_LOADERS[extension]);
}

function occtBase() {
    return (window.RESOURCEFIT_CONFIG && window.RESOURCEFIT_CONFIG.occtBase) || OCCT_BASE;
}

function loadOcct() {
    // The kernel is ~7 MB of wasm, so it is only fetched when a CAD file is opened.
    if (!window.occtimportjs) {
        return new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = `${occtBase()}occt-import-js.js`;
            script.onload = () => resolve(window.occtimportjs({ locateFile: (file) => occtBase() + file }));
            script.onerror = () => reject(new Error("The CAD geometry kernel could not be loaded."));
            document.head.appendChild(script);
        });
    }
    return window.occtimportjs({ locateFile: (file) => occtBase() + file });
}

function ensureNormals(geometry) {
    const normal = geometry.getAttribute("normal");
    if (!normal) {
        geometry.computeVertexNormals();
        return;
    }
    // Plenty of exporters write zero-length facet normals, which shade solid black.
    for (let index = 0; index < normal.count; index += 1) {
        if (normal.getX(index) || normal.getY(index) || normal.getZ(index)) return;
    }
    geometry.computeVertexNormals();
}

function geometryFromOcct(mesh) {
    const geometry = new THREE.BufferGeometry();
    geometry.setAttribute("position", new THREE.Float32BufferAttribute(mesh.attributes.position.array, 3));
    if (mesh.attributes.normal) {
        geometry.setAttribute("normal", new THREE.Float32BufferAttribute(mesh.attributes.normal.array, 3));
    }
    if (mesh.index) geometry.setIndex(new THREE.Uint32BufferAttribute(mesh.index.array, 1));
    ensureNormals(geometry);
    return geometry;
}

function material(color) {
    return new THREE.MeshStandardMaterial({
        color: color ? new THREE.Color(color[0], color[1], color[2]) : 0xb9c2cc,
        metalness: 0.12,
        roughness: 0.62,
        // Open shells and inconsistently wound STL facets are common in the wild.
        side: THREE.DoubleSide
    });
}

async function buildCadGroup(extension, buffer) {
    const occt = await loadOcct();
    const format = CAD_FORMATS[extension];
    const reader = { step: "ReadStepFile", iges: "ReadIgesFile", brep: "ReadBrepFile" }[format];
    const result = occt[reader](new Uint8Array(buffer), null);
    if (!result || !result.success) throw new Error("The CAD file could not be read.");
    if (!result.meshes.length) throw new Error("The CAD file contains no displayable geometry.");

    const group = new THREE.Group();
    result.meshes.forEach((mesh) => group.add(new THREE.Mesh(geometryFromOcct(mesh), material(mesh.color))));
    return group;
}

async function buildMeshGroup(extension, buffer, url) {
    const [modulePath, exportName] = MESH_LOADERS[extension];
    const loaderModule = await import(`three/addons/${modulePath}`);
    const loader = new loaderModule[exportName]();

    if (extension === "glb" || extension === "gltf") {
        const gltf = await loader.parseAsync(buffer, url.replace(/[^/]*$/, ""));
        return gltf.scene;
    }
    const parsed = loader.parse(extension === "obj" ? new TextDecoder().decode(buffer) : buffer);
    if (parsed.isBufferGeometry) {
        ensureNormals(parsed);
        return new THREE.Mesh(parsed, material());
    }
    parsed.traverse((child) => {
        if (!child.isMesh) return;
        ensureNormals(child.geometry);
        if (!child.material.map) child.material = material();
    });
    return parsed;
}

function countTriangles(root) {
    let triangles = 0;
    root.traverse((child) => {
        if (!child.isMesh) return;
        const geometry = child.geometry;
        triangles += (geometry.index ? geometry.index.count : geometry.getAttribute("position").count) / 3;
    });
    return Math.round(triangles);
}

export async function previewModel({ url, extension, viewport, onStage, onReady }) {
    const response = await fetch(url, { mode: "cors", credentials: "omit" });
    if (!response.ok) throw new Error(`The file download failed (HTTP ${response.status}).`);
    const buffer = await response.arrayBuffer();
    if (!buffer.byteLength) throw new Error("The downloaded file is empty.");

    const root = CAD_FORMATS[extension]
        ? await buildCadGroup(extension, buffer)
        : await buildMeshGroup(extension, buffer, url);

    // Downloading and meshing is the slow part; the caller keeps its spinner up
    // until here and only then reveals the viewport this renderer measures.
    if (onStage) onStage();

    const box = new THREE.Box3().setFromObject(root);
    if (box.isEmpty()) throw new Error("The model contains no displayable geometry.");
    const sphere = box.getBoundingSphere(new THREE.Sphere());
    const radius = sphere.radius || 1;

    const scene = new THREE.Scene();
    scene.background = new THREE.Color(0xeef1f4);
    scene.add(new THREE.HemisphereLight(0xffffff, 0x8d99ab, 2.2));
    const key = new THREE.DirectionalLight(0xffffff, 1.6);
    key.position.set(1, 1.4, 1.1);
    scene.add(key);
    const fill = new THREE.DirectionalLight(0xffffff, 0.7);
    fill.position.set(-1, -0.6, -0.8);
    scene.add(fill);
    scene.add(root);

    const renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
    viewport.replaceChildren(renderer.domElement);

    const camera = new THREE.PerspectiveCamera(45, 1, radius / 100, radius * 100);
    const direction = new THREE.Vector3(0.6, 0.45, 0.85).normalize();

    const controls = new OrbitControls(camera, renderer.domElement);
    controls.target.copy(sphere.center);
    // One-finger orbit, two-finger pinch zoom and pan come from OrbitControls itself.
    controls.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_PAN };

    function render() {
        renderer.render(scene, camera);
    }

    function frame() {
        // Fit on whichever field of view is narrower, so portrait embeds do not crop.
        const verticalFov = THREE.MathUtils.degToRad(camera.fov);
        const horizontalFov = 2 * Math.atan(Math.tan(verticalFov / 2) * camera.aspect);
        const distance = (radius / Math.sin(Math.min(verticalFov, horizontalFov) / 2)) * 1.15;
        camera.position.copy(sphere.center).addScaledVector(direction, distance);
        controls.target.copy(sphere.center);
        controls.update();
        render();
    }

    function resize() {
        const width = viewport.clientWidth || 1;
        const height = viewport.clientHeight || 1;
        renderer.setSize(width, height, false);
        camera.aspect = width / height;
        camera.updateProjectionMatrix();
        render();
    }

    // Damping is off, so rendering on demand keeps the embed from burning battery.
    controls.addEventListener("change", render);
    const observer = new ResizeObserver(resize);
    observer.observe(viewport);
    resize();
    frame();

    if (onReady) onReady({ triangles: countTriangles(root) });

    return {
        resetView: frame,
        dispose() {
            observer.disconnect();
            controls.dispose();
            renderer.dispose();
        }
    };
}
