<template>
   <q-layout view="hHh lpR fFf">

    <q-header elevated class="bg-bluemarine text-white" height-hint="98">
      <q-toolbar>
        <q-toolbar-title to="/">
          <q-avatar>
            <img src="../assets/isotipo-blanco.svg">
          </q-avatar>
          Formateador
        </q-toolbar-title>
      </q-toolbar>

      <q-tabs align="left">
        <q-route-tab to="/formateadores/esamur" label="ESAMUR" />
        <q-route-tab to="/formateadores/dcdc" label="DCDC" />
        <q-route-tab to="/formateadores/emasesa" label="EMASESA" />
      </q-tabs>
    </q-header>

    <q-page-container>
      <router-view />
    </q-page-container>

    <q-footer elevated class="bg-grey-8 text-white">
      <q-toolbar>
        <!-- <q-toolbar-title>
          <q-avatar>
            <img src="https://cdn.quasar.dev/logo-v2/svg/logo-mono-white.svg">
          </q-avatar>
          <div>Title</div>
        </q-toolbar-title> -->
      </q-toolbar>
    </q-footer>

  </q-layout>
</template>

<script>
export default {
  name: 'MainLayout'
}
</script>