<script lang="ts">
  export let compact = false;
</script>

<a class:compact class="brand" href="/" aria-label="InboxLink home">
  <span class="brand-mark" aria-hidden="true"><span></span></span>
  <span>InboxLink</span>
</a>

<style>
  .brand {
    align-items: center;
    color: var(--ink);
    display: inline-flex;
    font-size: 1.02rem;
    font-weight: 740;
    gap: 0.68rem;
    letter-spacing: -0.025em;
    text-decoration: none;
  }

  .brand-mark {
    align-items: center;
    background: linear-gradient(145deg, #7058ff, #3727ca);
    border-radius: 11px;
    box-shadow: 0 8px 20px rgba(67, 47, 211, 0.25);
    display: inline-flex;
    height: 30px;
    justify-content: center;
    transform: rotate(-6deg);
    width: 30px;
  }

  .brand-mark span {
    border: 2px solid white;
    border-left-width: 5px;
    border-radius: 3px 8px 8px 3px;
    height: 11px;
    width: 11px;
  }

  .compact .brand-mark {
    height: 26px;
    width: 26px;
  }
</style>
