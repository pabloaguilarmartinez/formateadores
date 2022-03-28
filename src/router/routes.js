const routes = [
  {
    path: '/',
    redirect: '/formateadores'
  },
  {
    path: '/formateadores',
    component: () => import('layouts/MainLayout.vue'),
    children: [
      { path: '', component: () => import('pages/Index.vue') },
      { path: 'dcdc', component: () => import('pages/FormateadorDcdc.vue') },
      { path: 'esamur', component: () => import('pages/FormateadorEsamur.vue') },
      { path: 'emasesa', component: () => import('pages/FormateadorEmasesa.vue') }
    ]
  },
  {
    path: '*',
    component: () => import('pages/Error404.vue')
  }
]

export default routes;