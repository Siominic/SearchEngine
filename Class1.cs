using Microsoft.AspNetCore.Builder;
using System;
using System.Threading;

namespace Search_EngineTLS;
public class RateLimiter
{
    int TokensPerSec;
    private int _tokens;
    private DateTime _lastRefill;
    private readonly object _lock = new object();

    public static void Rate_Limiter(int capacity, int TokensPerSec)
    {
        int _tokens = capacity;
        //make an if statement to check token quantity 
        if (_tokens > 0)
        {
            Program program = new Program();
            Program.Initialize();
        }
        else
        {
            // Start a background thread to refill tokens
            RefillTokens(capacity, _tokens);
        }
    }

    public static void RefillTokens(int _capacity, int _tokens)
    {
        int tokensPerSec = 2;
        while (true) //rewrite this so it doesn't loop, using maybe a do until for tokens to be refilled
        {
            Thread.Sleep(1000 / tokensPerSec);
            {
                do
                {
                    DateTime now = DateTime.Now;
                    DateTime _lastRefill = DateTime.Now;
                    double timePassed = (now - _lastRefill).TotalSeconds;
                    int tokensToAdd = (int)(timePassed * _capacity);
                    _tokens = Math.Min(_capacity, _tokens + tokensToAdd);
                    _lastRefill = now;
                } while (_tokens < 120);  //until the tokens are full


            }
        }
    }
}


