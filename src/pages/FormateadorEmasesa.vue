<template>
  <q-page class="flex justify-evenly items-center column">
    <q-card class="my-card" style="width: 400px">
      <q-card-section>
        <div class="text-h5">Asociar</div>
      </q-card-section>
      <q-separator inset />
      <q-card-section class="flex flex-center row">
        <q-input
          v-model.trim="areaActivo.val"
          label="Área Activo"
          class="flex q-px-md"
          style="width: 150px"
        />
        <q-input
          v-model.trim="codigoEstacion.val"
          label="Código Estación"
          class="flex q-px-md"
          style="width: 160px"
        />
        <q-file v-model="file.val" label="Seleccione fichero" style="width: 300px"/>
      </q-card-section>
      <q-card-actions>
        <q-space />
        <q-btn class="btn" label="Asociar" @click="loadFile()" />
      </q-card-actions>
    </q-card>
    <q-card class="my-card" style="width: 400px">
      <q-card-section>
        <div class="text-h5">Formatear</div>
      </q-card-section>
      <q-separator inset />
      <q-card-section class="text-h6 flex flex-center">
        A desarrollar
      </q-card-section>
    </q-card>
  </q-page>
</template>

<script>
import { associate } from "../service/EmasesaExcelService.js";

export default {
  name: "PageFormateadorEmasesa",
  data() {
    return {
      file: {
        val: null,
        isValid: true,
      },
      areaActivo: {
        val: '',
        isValid: true,
      },
      codigoEstacion: {
        val: '',
        isValid: true,
      },
      error: false,
    };
  },
  methods: {
    validateInputs() {
      if (this.file.val === null) {
        this.file.isValid = false;
      } else {
        this.file.isValid = true;
      }
      if (this.areaActivo.val === "") {
        this.areaActivo.isValid = false;
      } else {
        this.areaActivo.isValid = true;
      }
      if (this.codigoEstacion.val === "") {
        this.codigoEstacion.isValid = false;
      } else {
        this.codigoEstacion.isValid = true;
      }
    },
    loadFile() {
      this.validateInputs();
      if (
        !this.file.isValid ||
        !this.codigoEstacion.isValid ||
        !this.areaActivo.isValid
      ) {
        // this.error = true;
        this.$q.notify({
          type: "warning",
          message: `Rellene todos los campos.`,
          actions: [{ icon: "close", color: "black" }],
        });
      } else {
        try {
          associate(
            this.file.val,
            this.codigoEstacion.val,
            this.areaActivo.val
          );
        } catch (exception) {
          console.log(exception);
        }
      }
    },
  },
};
</script>
