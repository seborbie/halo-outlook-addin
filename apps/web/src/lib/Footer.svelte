<script lang="ts">
  import Brand from "$lib/Brand.svelte";
</script>

<footer class="footer">
  <div class="shell footer-grid">
    <div>
      <Brand compact />
      <p class="footer-intro">A calmer bridge between Outlook and HaloPSA.</p>
    </div>
    <div class="footer-links">
      <a href="/pricing">Pricing</a>
      <a href="/signup">Sign up</a>
      <a href="/docs/overview">Documentation</a>
      <a href="/docs/security">Security</a>
    </div>
  </div>
  <div class="shell legal-row">
    <p>© {new Date().getFullYear()} InboxLink.</p>
    <p>
      HaloPSA, the HaloPSA name, and Halo logos are trademarks of Halo Service Solutions.
      InboxLink is an independent product and is not affiliated with or endorsed by Halo.
    </p>
  </div>
</footer>
