<template>
  <q-page class="flex flex-center">
    <q-dialog v-model="error">
      <q-card style="width: 300px">
        <q-card-section>
          <div class="text-h6">Error</div>
        </q-card-section>

        <q-card-section class="q-pt-none">
          Ha habido un error. Revise que ha rellenado todos los campos.
        </q-card-section>

        <q-card-actions align="right" class="bg-white text-teal">
          <q-btn flat label="OK" v-close-popup />
        </q-card-actions>
      </q-card>
    </q-dialog>
    <q-input
      v-model.trim="nombreEdar.val"
      label="Nombre EDAR"
      class="flex q-px-md"
    />
    <q-input
      v-model.trim="identificador.val"
      label="Identificador"
      class="flex q-px-md"
    />
    <q-file
      v-model="file.val"
      label="Seleccione fichero"
      class="flex q-px-md"
    />
    <q-btn
      class="bg-bluemarine"
      color="text-white"
      label="Formatear"
      @click="loadFile()"
    />
  </q-page>
</template>

<script>
import { format } from "../service/EsamurExcelService.js";

export default {
  name: "PageFormateadorEsamur",
  data() {
    return {
      file: {
        val: null,
        isValid: true,
      },
      nombreEdar: {
        val: "",
        isValid: true,
      },
      identificador: {
        val: "",
        isValid: true,
      },
      error: false,
    };
  },
  methods: {
    validateInputs() {
      if (this.nombreEdar.val === "") {
        this.nombreEdar.isValid = false;
      }
      if (this.identificador.val === "") {
        this.identificador.isValid = false;
      }
      if (this.file.val === null) {
        this.file.isValid = false;
      }
    },
    loadFile() {
      this.validateInputs();
      if (
        !this.file.isValid ||
        !this.identificador.isValid ||
        !this.nombreEdar.isValid
      ) {
        this.error = true;
      } else {
        try {
          format(this.file.val, this.nombreEdar.val, this.identificador.val);
        } catch (exception) {
          console.log(exception);
        }
      }
    },
  },
};
</script>