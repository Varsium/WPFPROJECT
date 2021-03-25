using System.Collections.Generic;
using EindOpdrachtLeandro.model;

namespace EindOpdrachtLeandro.Repositories
{
    interface Isportclub
    {
        IList<Sportclub> GetAll();
        IList<string> GetSportclubNaam();
    }

}
