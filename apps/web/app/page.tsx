"use client";

import { type FormEvent, useEffect, useRef } from "react";
import * as THREE from "three";

const phoneNumbers = ["+91-9414176451", "+91-9828516951"];
const whatsappNumber = "919414176451";
const mapsUrl = "https://maps.app.goo.gl/Jfg2m6eGyQuyBQC97";

const products = [
  {
    name: "Wheat",
    season: "Rabi",
    description:
      "High-vigour wheat seed options selected for strong tillering, uniform maturity, and dependable farmer performance."
  },
  {
    name: "Soyabean",
    season: "Kharif",
    description:
      "Quality soyabean seed portfolio focused on plant stand, pod setting, and resilient field establishment."
  },
  {
    name: "Paddy",
    season: "Kharif",
    description:
      "Paddy seed selections for consistent germination, healthy crop growth, and region-ready cultivation needs."
  },
  {
    name: "Chana",
    season: "Rabi",
    description:
      "Chana seed lines developed for strong emergence, sturdy crop habit, and reliable pulse production."
  },
  {
    name: "Moong",
    season: "Zaid / Kharif",
    description:
      "Moong seed options designed for quick crop cycles, uniform growth, and practical farmer profitability."
  },
  {
    name: "Mustard",
    season: "Rabi",
    description:
      "Mustard seed portfolio oriented toward vigorous plants, balanced branching, and dependable oilseed output."
  }
];

const qualitySteps = [
  "Seed selection from trusted production channels",
  "Field-level monitoring and variety-wise traceability",
  "Processing, grading, and lot discipline",
  "Farmer feedback loops across Rajasthan and nearby markets"
];

function SeedScene() {
  const mountRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const mount = mountRef.current;
    if (!mount) return;

    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(42, 1, 0.1, 100);
    camera.position.set(0, 2.4, 7.2);

    const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.setClearColor(0xffffff, 0);
    mount.appendChild(renderer.domElement);

    const keyLight = new THREE.DirectionalLight(0xffffff, 2.1);
    keyLight.position.set(3, 4, 5);
    scene.add(keyLight);
    scene.add(new THREE.AmbientLight(0xf8ffe2, 1.1));

    const fieldGroup = new THREE.Group();
    const soilMaterial = new THREE.MeshStandardMaterial({
      color: 0xd6b45d,
      roughness: 0.85,
      metalness: 0.04
    });
    const greenMaterial = new THREE.MeshStandardMaterial({
      color: 0x75b843,
      roughness: 0.72,
      metalness: 0.03
    });

    for (let i = 0; i < 9; i += 1) {
      const band = new THREE.Mesh(
        new THREE.BoxGeometry(8.5, 0.08, 0.18),
        i % 2 === 0 ? greenMaterial : soilMaterial
      );
      band.position.set(0, -1.35 + i * 0.18, -1.9 + i * 0.23);
      band.rotation.x = -0.58;
      band.rotation.z = i % 2 === 0 ? 0.02 : -0.02;
      fieldGroup.add(band);
    }
    scene.add(fieldGroup);

    const seedGroup = new THREE.Group();
    const seedMaterial = new THREE.MeshPhysicalMaterial({
      color: 0xf5c34c,
      roughness: 0.38,
      metalness: 0.08,
      clearcoat: 0.35
    });
    const leafMaterial = new THREE.MeshStandardMaterial({
      color: 0x2e9d52,
      roughness: 0.45
    });

    for (let i = 0; i < 18; i += 1) {
      const seed = new THREE.Mesh(new THREE.SphereGeometry(0.12, 28, 20), seedMaterial);
      const angle = (i / 18) * Math.PI * 2;
      const radius = 1.15 + (i % 3) * 0.28;
      seed.scale.set(1, 0.72, 0.9);
      seed.position.set(Math.cos(angle) * radius, Math.sin(angle * 1.4) * 0.42, Math.sin(angle) * 0.9);
      seed.rotation.set(angle * 0.2, angle, angle * 0.1);
      seedGroup.add(seed);
    }

    const stem = new THREE.Mesh(
      new THREE.CylinderGeometry(0.025, 0.045, 1.55, 18),
      new THREE.MeshStandardMaterial({ color: 0x3a8d45, roughness: 0.52 })
    );
    stem.position.y = 0.02;
    stem.rotation.z = -0.13;
    seedGroup.add(stem);

    const leafOne = new THREE.Mesh(new THREE.SphereGeometry(0.34, 32, 18), leafMaterial);
    leafOne.scale.set(1.45, 0.2, 0.55);
    leafOne.position.set(0.33, 0.72, 0);
    leafOne.rotation.z = -0.42;
    seedGroup.add(leafOne);

    const leafTwo = leafOne.clone();
    leafTwo.position.set(-0.34, 0.52, 0.02);
    leafTwo.rotation.z = 0.55;
    seedGroup.add(leafTwo);

    seedGroup.position.set(0, 0.45, 0);
    scene.add(seedGroup);

    const resize = () => {
      const { width, height } = mount.getBoundingClientRect();
      renderer.setSize(width, height);
      camera.aspect = width / Math.max(height, 1);
      camera.updateProjectionMatrix();
    };
    resize();

    let frame = 0;
    let animationId = 0;
    const animate = () => {
      frame += 0.01;
      seedGroup.rotation.y = frame * 0.55;
      seedGroup.position.y = 0.45 + Math.sin(frame * 1.4) * 0.08;
      fieldGroup.rotation.y = Math.sin(frame * 0.45) * 0.08;
      camera.position.x = Math.sin(frame * 0.32) * 0.35;
      camera.lookAt(0, 0.05, 0);
      renderer.render(scene, camera);
      animationId = window.requestAnimationFrame(animate);
    };
    animate();

    window.addEventListener("resize", resize);
    return () => {
      window.removeEventListener("resize", resize);
      window.cancelAnimationFrame(animationId);
      renderer.dispose();
      mount.removeChild(renderer.domElement);
    };
  }, []);

  return <div className="site-hero-scene" ref={mountRef} aria-hidden="true" />;
}

export default function HomePage() {
  const handleEnquirySubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const message = [
      "Hello Krishiv Seeds, I want to enquire about your seed products.",
      `Name: ${formData.get("name") || ""}`,
      `Phone: ${formData.get("phone") || ""}`,
      `Interest: ${formData.get("interest") || ""}`,
      `Message: ${formData.get("message") || ""}`
    ].join("\n");

    window.location.href = `https://wa.me/${whatsappNumber}?text=${encodeURIComponent(message)}`;
  };

  return (
    <main className="site-shell">
      <nav className="site-nav" aria-label="Primary navigation">
        <a className="site-brand" href="#home" aria-label="Krishiv home">
          <img src="/krishiv-logo.svg" alt="Krishiv Seeds" />
        </a>
        <div className="site-nav-links">
          <a href="#products">Products</a>
          <a href="#quality">Quality</a>
          <a href="#dealers">Dealers</a>
          <a href="#contact">Contact</a>
        </div>
        <a className="site-nav-cta" href={`https://wa.me/${whatsappNumber}`}>WhatsApp</a>
      </nav>

      <section className="site-hero" id="home">
        <div className="site-hero-copy">
          <p className="site-kicker">Krishiv Agri Genetics LLP</p>
          <h1>Quality seeds for confident harvests.</h1>
          <p className="site-lede">
            Founded in Kota by Jambu Jain and Piyush Jain, Krishiv Seeds delivers trusted seed solutions for farmers,
            dealers, and distributors with a growing network of 10,000+ satisfied farmers.
          </p>
          <div className="site-actions">
            <a className="site-button primary" href={`tel:${phoneNumbers[0]}`}>Call now</a>
            <a className="site-button secondary" href="#products">Browse seeds</a>
          </div>
          <div className="site-metrics" aria-label="Company highlights">
            <span><strong>2023</strong> Founded</span>
            <span><strong>10,000+</strong> Farmers served</span>
            <span><strong>6</strong> Core crops</span>
          </div>
        </div>
        <SeedScene />
      </section>

      <section className="site-band">
        <p>Our mission is to bring reliable genetics, disciplined seed quality, and practical field support to every acre we serve.</p>
      </section>

      <section className="site-section" id="products">
        <div className="site-section-head">
          <p className="site-kicker">Seed Portfolio</p>
          <h2>Built for Rajasthan's farming rhythm</h2>
          <p>Image spaces are ready. When you share product photos or packet images, they can be dropped straight into these cards.</p>
        </div>
        <div className="site-product-grid">
          {products.map((product, index) => (
            <article className="site-product-card" key={product.name}>
              <div className="site-product-visual">
                <span>{product.name.slice(0, 1)}</span>
                <i>{product.season}</i>
              </div>
              <div>
                <p className="site-product-index">0{index + 1}</p>
                <h3>{product.name}</h3>
                <p>{product.description}</p>
              </div>
            </article>
          ))}
        </div>
      </section>

      <section className="site-split" id="quality">
        <div>
          <p className="site-kicker">Research & Quality</p>
          <h2>Sharp process, grounded field sense.</h2>
          <p>
            Krishiv Seeds is young by year and serious by practice. The company focuses on dependable seed lots,
            clean processing discipline, and product choices that make sense for real farm conditions.
          </p>
        </div>
        <div className="site-quality-panel">
          {qualitySteps.map((step) => (
            <div className="site-quality-row" key={step}>
              <span />
              <p>{step}</p>
            </div>
          ))}
        </div>
      </section>

      <section className="site-dealer" id="dealers">
        <div>
          <p className="site-kicker">Dealer & Distributor Enquiry</p>
          <h2>Partner with Krishiv Seeds</h2>
          <p>
            We welcome dealers, distributors, and farmer groups looking for a responsive seed partner with focused crop
            categories and clear communication.
          </p>
        </div>
        <form className="site-form" onSubmit={handleEnquirySubmit}>
          <label>
            Name
            <input type="text" name="name" placeholder="Your name" />
          </label>
          <label>
            Phone
            <input type="tel" name="phone" placeholder="+91" />
          </label>
          <label>
            Interest
            <select name="interest" defaultValue="">
              <option value="" disabled>Select enquiry type</option>
              <option>Dealer enquiry</option>
              <option>Distributor enquiry</option>
              <option>Farmer product enquiry</option>
            </select>
          </label>
          <label>
            Message
            <textarea name="message" placeholder="Tell us your crop, district, or requirement" />
          </label>
          <button className="site-button primary" type="submit">
            Send enquiry on WhatsApp
          </button>
        </form>
      </section>

      <section className="site-contact" id="contact">
        <div>
          <p className="site-kicker">Contact</p>
          <h2>Krishiv Agri Genetics LLP</h2>
          <p>G-374, I.P.I.A., Road No. 7, Kota, Rajasthan</p>
        </div>
        <div className="site-contact-grid">
          <a href={`tel:${phoneNumbers[0]}`}>{phoneNumbers[0]}</a>
          <a href={`tel:${phoneNumbers[1]}`}>{phoneNumbers[1]}</a>
          <a href="mailto:piyush.industries.kota@gmail.com">piyush.industries.kota@gmail.com</a>
          <a href={mapsUrl} target="_blank" rel="noreferrer">Open Google Maps</a>
        </div>
      </section>
    </main>
  );
}
