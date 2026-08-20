<script lang="ts">
  import Brand from "$lib/Brand.svelte";

  let menuOpen = false;
</script>

<header class="site-header">
  <nav class="shell nav" aria-label="Main navigation">
    <Brand />
    <button
      class="menu-button"
      type="button"
      aria-label="Toggle navigation"
      aria-expanded={menuOpen}
      onclick={() => (menuOpen = !menuOpen)}
    >
      <span></span><span></span>
    </button>
    <div class:open={menuOpen} class="nav-links">
      <a href="/#features" onclick={() => (menuOpen = false)}>Features</a>
      <a href="/pricing" onclick={() => (menuOpen = false)}>Pricing</a>
      <a href="/docs/overview" onclick={() => (menuOpen = false)}>Docs</a>
      <a class="button button-small" href="/signup" onclick={() => (menuOpen = false)}>
        Start a pilot <span aria-hidden="true">↗</span>
      </a>
    </div>
  </nav>
</header>
