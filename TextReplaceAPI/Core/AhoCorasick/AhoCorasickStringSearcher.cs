/******************************************************************
 * Copyright (c) QuickReplace, LLC. All Rights Reserved.
 * 
 * This file is part of the QuickReplace Developer Library. You may not
 * use this file for commercial or business use except in compliance with
 * the QuickReplace license. You may obtain a copy of the license at:
 * 
 *   https://www.quickreplace.io/pricing
 * 
 * You may view the terms of this license at:
 * 
 *   https://www.quickreplace.io/eula
 * 
 * This file or any others in the QuickReplace Developer Library may not
 * be modified, shared, or distributed outside of the terms of the license
 * agreement by organizations or individuals who are covered by a license.
 * 
 *****************************************************************/

namespace TextReplaceAPI.Core.AhoCorasick
{
    internal class AhoCorasickStringSearcher
    {
        // GotoTransations:
        // For each state, we have a Dictionary<char, newState>
        private readonly Dictionary<int, Dictionary<string, int>> GotoTransitions =
            new Dictionary<int, Dictionary<string, int>>();

        // FailTransistions:
        // Dictionary<state, failState>
        private readonly Dictionary<int, int> FailTransitions = new Dictionary<int, int>();

        // Output:
        // Dictionary<state, (list of outputs)>
        private readonly Dictionary<int, HashSet<string>> Output = new Dictionary<int, HashSet<string>>();

        // Total number of states in automaton
        public int NumStates { get; private set; }

        // Total number of outputs added to nodes
        public long TotalOutputs { get; private set; }

        public bool CaseSensitive { get; private set; }

        private readonly int StartState;

        private readonly bool debugMode = false;


        public AhoCorasickStringSearcher(bool caseSensitive)
        {
            CaseSensitive = caseSensitive;
            StartState = CreateNewState();
        }

        private int CreateNewState()
        {
            int state = NumStates++;
            if (CaseSensitive)
            {
                GotoTransitions.Add(state, new Dictionary<string, int>());
            }
            else
            {
                GotoTransitions.Add(state, new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase));
            }
            
            return state;
        }

        private Dictionary<string, int> GetStateTransitions(int state)
        {
            Dictionary<string, int>? transitions = (CaseSensitive) ?
                new Dictionary<string, int>() :
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (!GotoTransitions.TryGetValue(state, out transitions))
            {
                throw new ApplicationException(string.Format("State {0} is not defined.", state));
            }
            return transitions;
        }

        private int Goto(int state, string c)
        {
            var transitions = GetStateTransitions(state);
            int newState;
            if (!transitions.TryGetValue(c, out newState))
            {
                return -1;
            }
            return newState;
        }

        private void AddStateTransition(int state, string c, int newState)
        {
            var transitions = GetStateTransitions(state);
            transitions.Add(c, newState);
        }

        private HashSet<string>? GetStateOutputs(int state)
        {
            HashSet<string>? outputs;
            if (Output.TryGetValue(state, out outputs))
            {
                return outputs;
            }
            return null;
        }

        private void AddOutput(int state, string outputKey)
        {
            HashSet<string>? outputs;
            if (!Output.TryGetValue(state, out outputs))
            {
                outputs = new HashSet<string>();
                Output.Add(state, outputs);
            }
            if (outputs.Add(outputKey))
            {
                ++TotalOutputs;
            }
        }

        private void AddOutputs(int state, HashSet<string> outputsToAdd)
        {
            HashSet<string>? outputs;
            if (!Output.TryGetValue(state, out outputs))
            {
                Output.Add(state, outputsToAdd);
                TotalOutputs += outputsToAdd.Count;
            }
            else
            {
                int count = outputs.Count;
                foreach (var o in outputsToAdd)
                {
                    outputs.Add(o);
                }
                TotalOutputs += (outputs.Count - count);
            }
        }

        public void AddItem(string item)
        {
            // Now add to the automaton
            int state = StartState;
            foreach (var c in item)
            {
                int newState = Goto(state, c.ToString());
                if (newState == -1)
                {
                    newState = CreateNewState();
                    AddStateTransition(state, c.ToString(), newState);
                }
                state = newState;
            }
            // The tree has been updated for this word
            // Add the out transition
            AddOutput(state, item);
        }

        private int SearchGoto(int state, string c)
        {
            int newState = Goto(state, c);
            if (newState == -1 && state == StartState)
            {
                return StartState;
            }
            return newState;
        }

        public void CreateFailureFunction()
        {
            // only used for debug mode
            // initialized at 1 because we check the start state separately.
            int statesChecked = 1;

            FailTransitions[StartState] = StartState;
            Queue<int> StateQueue = new Queue<int>();
            var transitions = GetStateTransitions(StartState);
            foreach (var kvp in transitions)
            {
                StateQueue.Enqueue(kvp.Value);
                FailTransitions.Add(kvp.Value, StartState);
            }
            while (StateQueue.Count > 0)
            {
                int r = StateQueue.Dequeue();
                transitions = GetStateTransitions(r);
                foreach (var kvp in transitions)
                {
                    StateQueue.Enqueue(kvp.Value);
                    int state = FailTransitions[r];
                    while (SearchGoto(state, kvp.Key) == -1)
                    {
                        state = FailTransitions[state];
                    }
                    int failState = SearchGoto(state, kvp.Key);
                    FailTransitions[kvp.Value] = failState;
                    // now merge outputs from the fail state to the transition state
                    var failOutputs = GetStateOutputs(failState);
                    if (failOutputs != null)
                    {
                        AddOutputs(kvp.Value, failOutputs);
                        // This should get rid of excess, at the cost of some processor time.
                        //failOutputs.TrimExcess();
                    }
                }

                if (debugMode)
                {
                    if ((++statesChecked % 10000) == 0)
                    {
                        Console.Write("\r{0:N0} checked.\t{1:N0} in queue.\t{2:N0} outputs", statesChecked, StateQueue.Count, TotalOutputs);
                    }
                }
            }

            if (debugMode)
            {
                Console.WriteLine();
                Console.WriteLine("{0:N0} states checked.\t{1:N0} outputs", statesChecked, TotalOutputs);
            }

        }

        /// <summary>
        /// Searches a text for matches of some specific words using the Aho-Corsasick algorithm
        /// </summary>
        /// <param name="text"></param>
        /// <returns>
        /// An enumerable containing StringMatch objects.
        /// These contain the text found and the position it was found at.
        /// </returns>
        public IEnumerable<StringMatch> Search(string text)
        {
            var foundItems = new List<StringMatch>();
            int state = StartState;
            //foreach (var c in text)
            for (int pos = 0; pos < text.Length; ++pos)
            {
                string c = text[pos].ToString();
                while (Goto(state, c) == -1)
                {
                    if (state == StartState)
                    {
                        goto EndWord;
                    }
                    state = FailTransitions[state];
                }
                state = Goto(state, c);
                var outputs = GetStateOutputs(state);
                if (outputs != null)
                {
                    foreach (var o in outputs)
                    {
                        foundItems.Add(new StringMatch(o, pos - o.Length + 1));
                    }
                }
            EndWord:
                ;
            }
            return foundItems;
        }
    }
}
