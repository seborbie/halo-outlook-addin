<script lang="ts">
  import "../app.css";
  import Footer from "$lib/Footer.svelte";
  import Nav from "$lib/Nav.svelte";

  let { children } = $props();
</script>

<svelte:head>
  <title>InboxLink — Outlook meets HaloPSA</title>
  <meta name="description" content="Attach Outlook conversations to HaloPSA tickets without breaking your flow." />
</svelte:head>

<Nav />
<main>{@render children()}</main>
<Footer />
