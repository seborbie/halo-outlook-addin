<script lang="ts">
  import { docsBySection } from "$lib/docs";
</script>

<svelte:head>
  <title>Documentation — InboxLink</title>
  <meta name="description" content="Setup, deploy, secure, and troubleshoot InboxLink for Outlook and HaloPSA." />
</svelte:head>

<section class="shell docs-home">
  <div class="docs-home-header"><p class="eyebrow">Documentation</p><h1>From first deploy to daily operation.</h1><p class="lede">Practical guides for Microsoft admins, HaloPSA admins, and the people operating InboxLink.</p></div>
  {#each Object.entries(docsBySection) as [section, entries]}
    <div class="docs-section"><h2>{section}</h2><div class="docs-card-grid">
      {#each entries as doc}<a class="doc-card" href={`/docs/${doc.slug}`}><h3>{doc.title} <span aria-hidden="true">→</span></h3><p>{doc.description}</p></a>{/each}
    </div></div>
  {/each}
</section>
