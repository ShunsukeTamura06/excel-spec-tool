<template>
  <UApp>
    <!-- 画面遷移中であることを常に視認できるよう、トップに細い進捗バーを出す.
         vue-router の before/after フックに自動で繋がっており、ページ間の
         await useAsyncData 等が走っている間は伸び続ける. -->
    <NuxtLoadingIndicator color="var(--ui-color-primary-500)" :height="3" />
    <NuxtLayout>
      <NuxtPage />
    </NuxtLayout>
  </UApp>
</template>
