import { useEffect, useMemo } from "react";
import RepoLinkCard from "@/components/RepoLinkCard";

const REPO_URL = "https://github.com/fwehiu/Calculator_fruit";

const Index = () => {
  useEffect(() => {
    document.title = "Calculator Fruit – GitHub Link";

    const ensure = (selector: string, create: () => HTMLElement) => {
      let el = document.querySelector(selector) as HTMLElement | null;
      if (!el) {
        el = create();
        document.head.appendChild(el);
      }
      return el;
    };

    const desc = ensure('meta[name="description"]', () => {
      const m = document.createElement("meta");
      m.setAttribute("name", "description");
      return m;
    }) as HTMLMetaElement;
    desc.setAttribute("content", "Quick link to the Calculator Fruit GitHub repository.");

    const canonical = ensure('link[rel="canonical"]', () => {
      const l = document.createElement("link");
      l.setAttribute("rel", "canonical");
      return l;
    }) as HTMLLinkElement;
    canonical.setAttribute("href", window.location.href);
  }, []);

  const jsonLd = useMemo(() => ({
    "@context": "https://schema.org",
    "@type": "SoftwareSourceCode",
    name: "Calculator Fruit",
    codeRepository: REPO_URL,
    programmingLanguage: "JavaScript",
    url: window.location.href,
  }), []);

  return (
    <>
      <main className="min-h-screen flex items-center justify-center bg-background">
        <section className="w-full max-w-3xl mx-auto px-6 py-16 rounded-lg border bg-gradient-to-b from-primary/5 to-accent/5">
          <header className="text-center mb-10">
            <h1 className="text-4xl font-bold tracking-tight mb-3">Calculator Fruit GitHub Repository</h1>
            <p className="text-lg text-muted-foreground">Access the source code and details via the official GitHub repo.</p>
          </header>

          <RepoLinkCard href={REPO_URL} />
        </section>
      </main>
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />
    </>
  );
};

export default Index;
