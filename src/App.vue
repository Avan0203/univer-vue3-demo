<template>
  <div id="app">
    <div>
      <button @click="getData" title="print current workbook data to console">Get Data</button>
      <button @click="destroy">destroy</button>
      <button @click="show">show</button>
      <input
        ref="fileInputRef"
        type="file"
        accept=".xlsx"
        style="display: none"
        @change="onFileChange"
      />
      <button type="button" @click="triggerImport">导入 xlsx</button>
      <button type="button" @click="exportXlsx">导出 xlsx</button>
    </div>
    <UniverSheet v-if="isShow" id="sheet" ref="univerRef" :data="data" @import="data = $event" />
  </div>
</template>

<script setup lang="ts">
import UniverSheet from './components/UniverSheet.vue'
import { DEFAULT_WORKBOOK_DATA } from './assets/default-workbook-data'
import { ref } from 'vue';

const data = ref(DEFAULT_WORKBOOK_DATA);
const univerRef = ref<InstanceType<typeof UniverSheet> | null>(null);
const fileInputRef = ref<HTMLInputElement | null>(null);
const isShow = ref(true);

const getData = () => {
  const result = univerRef.value?.getData();
  console.log(JSON.stringify(result, null, 2));
};

const destroy = () => {
  isShow.value = false;
};

const show = () => {
  isShow.value = true;
};

const triggerImport = () => {
  fileInputRef.value?.click();
};

const onFileChange = async (e: Event) => {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;
  await univerRef.value?.importFromFile(file);
  input.value = '';
};

const exportXlsx = () => {
  univerRef.value?.exportToFile('univer');
};

</script>

<style>
html,
body {
  margin: 0;
  padding: 0;
}

#app {
  height: 100vh;
  display: flex;
  flex-direction: column;
}

#sheet {
  /** The height of the Union uses the height of the parent container */
  flex: 1;
}
</style>
