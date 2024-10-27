
namespace importacionBono.Components.Model;

public class BonoReaglo
{
    public string BonoRegalo { get; set; } = default!;
    public double Valor { get; set; } = default!;
    public double Porcentaje { get; set; } = default!;

    public void Clear()
    {
        BonoRegalo = default!;
        Valor = default!;
        Porcentaje = default!;
    }
}