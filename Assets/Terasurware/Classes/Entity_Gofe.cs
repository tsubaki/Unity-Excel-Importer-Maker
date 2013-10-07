using UnityEngine;
using System.Collections;
using System.Collections.Generic;

public class Entity_Gofe : ScriptableObject
{
	
	public List<Param> list = new List<Param> ();
	
	[System.SerializableAttribute]
	public class Param
	{
		
		public string skill;
		public string effect;
		public double damage;
		public bool isEnable;
	}
}

