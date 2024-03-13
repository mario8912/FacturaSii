using System;

namespace Entidades.utils
{
    public class EventoProgreso
    {
        private int _progreso = 0;
        public int ValorMaximoBarraProgreso { get; set; }

        public event EventHandler<int> ProgresoCambiado;

        public void AumentarProgreso()
        {
            _progreso++;
            OnProgresoCambiado();
        }

        protected virtual void OnProgresoCambiado()
        {
            ProgresoCambiado?.Invoke(this, _progreso);
        }
    }
}

