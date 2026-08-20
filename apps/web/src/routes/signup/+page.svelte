<script lang="ts">
  import SignupForm from "$lib/SignupForm.svelte";
</script>

<svelte:head>
  <title>Start a pilot — InboxLink</title>
  <meta name="description" content="Create an isolated InboxLink organisation and connect it to your HaloPSA instance." />
</svelte:head>

<section class="shell signup-shell">
  <div class="signup-copy">
    <p class="eyebrow">Start a 14-day pilot</p>
    <h1>Your inbox can be part of the service desk.</h1>
    <p class="lede">Create your company workspace, connect your own HaloPSA OAuth application, then deploy the add-in to a pilot group.</p>
    <ul class="mini-list"><li>No payment card</li><li>Microsoft-verified company identity</li><li>Hard tenant isolation from day one</li></ul>
    <p class="lede" style="font-size:.82rem">By continuing, you acknowledge that InboxLink is an independent product. HaloPSA, the HaloPSA name, and Halo logos are trademarks of Halo Service Solutions.</p>
  </div>
  <div class="signup-card">
    <h2>Create your company workspace</h2>
    <p>You’ll sign in with a Microsoft account from the organisation you are registering.</p>
    <SignupForm />
  </div>
</section>
