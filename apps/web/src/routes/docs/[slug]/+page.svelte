<script lang="ts">
  import { page } from "$app/state";
  import { docs, docsBySection } from "$lib/docs";

  const modules = import.meta.glob("../../../content/docs/*.md", { eager: true }) as Record<
    string,
    { default: unknown; metadata?: Record<string, string> }
  >;
  const slug = $derived(page.params.slug);
  const entry = $derived(docs.find((doc) => doc.slug === slug));
  const module = $derived(modules[`../../../content/docs/${slug}.md`]);
</script>

<svelte:head>
  <title>{entry ? `${entry.title} — InboxLink docs` : "Documentation — InboxLink"}</title>
  <meta name="description" content={entry?.description || "InboxLink documentation"} />
</svelte:head>

<div class="shell docs-layout">
  <aside class="docs-sidebar" aria-label="Documentation navigation">
    {#each Object.entries(docsBySection) as [section, entries]}<section><h2>{section}</h2>{#each entries as doc}<a class:active={doc.slug === slug} href={`/docs/${doc.slug}`}>{doc.title}</a>{/each}</section>{/each}
  </aside>
  <article class="doc-content">
    {#if entry && module}
      <p class="doc-kicker">{entry.section}</p>
      <p class="doc-description">{entry.description}</p>
      {@const Content = module.default as any}
      <Content />
    {:else}
      <h1>Guide not found</h1><p>Return to the <a href="/docs">documentation index</a>.</p>
    {/if}
  </article>
</div>
