
from django.views.generic import FormView
from django.core.urlresolvers import reverse_lazy
from .forms import UserForm
from .models import cat000usuario

class Registro(FormView):
    template_name= 'webportal/registro_usuario.html'
    form_class = UserForm
    success_url = reverse_lazy('registro')
    
    def form_valid(self, form):
        user = form.save()
        usuario_ = cat000usuario()
        usuario_.usuario = user
        usuario_.departamento = form.cleaned_data['departamento']
        usuario_.save()
        
        return super(Registro, self).form_valid(form)
    
