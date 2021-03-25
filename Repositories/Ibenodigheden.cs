using System.Collections.Generic;
using EindOpdrachtLeandro.model;

namespace EindOpdrachtLeandro.Repositories
{
    interface Ibenodigheden
    {
        IList<Benodigheden> GetAll();

    }
}
