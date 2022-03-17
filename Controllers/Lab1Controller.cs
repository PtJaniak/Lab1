using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using System;



namespace Lab1.Controllers
{


[Route("api/[controller]")]
[ApiController]

public class Lab1Controller : ControllerBase
{
DataContainer dt;
        private DataContainer _data;

        public Lab1Controller(DataContainer d)
{
    dt = d;
}
[HttpGet("rekordydlakraju/")]
public List<Dane> RekordyDlaKraju([FromQuery] string country) 
{
  return dt.NaszeDane.Where(x => x.Country == country).ToList();
}
[HttpGet("rekordydlasegmentu/")]
public List<Dane> RekordyDlaSegmentu([FromQuery] string segment) 
{
  return dt.NaszeDane.Where(x => x.Segment == segment).ToList();
}
[HttpGet("rekordydlaproduktu/")]
public List<Dane> RekordyDlaProduktu([FromQuery] string product) 
{
  return dt.NaszeDane.Where(x => x.Product == product).ToList();
}

[HttpGet("raport/")]

public IEnumerable<Raport> Raport()
        {
            var raport = new List<Raport>();

            var groups = dt.NaszeDane.GroupBy(s => new { s.Country, s.Segment })
                .OrderBy(g => g.Key.Country)
                .ThenBy(g => g.Key.Segment);

            foreach (var group in groups)
            {
                raport.Add(new Raport
                {
                    Country = group.Key.Country,
                    Segment = group.Key.Segment,
                    UnitsSold = group.Sum(s => s.UnitsSold)
                });
            }

            return raport;
        }

[HttpPost("DodajWpis")]
public void AddData(Dane newDane)
{
  dt.AddDataToExcel(newDane);
  
}

[HttpDelete("UsunWpis")]
public bool RemoveData(int id)
        {
            return dt.RemoveDataFromExcel(id);
           
        }

[HttpGet("WyswielWpis/")]
public List<Dane> Wpis([FromQuery] int id) 
{
  return dt.NaszeDane.Where(x => x.Id == id).ToList();
}
[HttpGet("ilerekordow/")]
public string IleRekordow()
{
    return $"Wczytano {dt.NaszeDane.Count()} rekordów";
}
}
}


    

    
        
        
    

