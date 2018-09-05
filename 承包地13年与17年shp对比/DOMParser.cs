using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace 承包地13年与17年shp对比
{
    public class DOMParser
    {
        Regex reg_att = new Regex(@"(\w+)\s*=\s*\""(.*?)\""");
        public DOMNode parse_text(string text)
        {
            DOMNode root = new DOMNode();
            root.Nodes = new List<DOMNode>();
            int index = 0;
            while (index < text.Length - 1)
            {
                DOMNode node = parse_next_node(text, ref index);
                root.Nodes.Add(node);
            }
            return root;
        }

        DOMNode parse_next_node(string text, ref int index)
        {
            StringBuilder node_outer_text = new StringBuilder();
            StringBuilder node_inner_text = new StringBuilder();
            bool node_begin = false;
            DOMNode node = null;
            while (index < text.Length - 1)
            {
                if (text[index] == '<' && (index < text.Length - 1) && text[index + 1] == '/') //标签关闭
                {
                    string node_name = get_next_token(text, '>', index + 2);
                    if (!node_begin)
                    {
                        node_outer_text.AppendFormat("</{0}>", node_name);
                        node_inner_text.AppendFormat("</{0}>",node_name);
                        node = new DOMNode()
                        {
                            Name = "text",
                            IsClosed = true,
                        };
                        index += node_name.Length + 3;
                    }
                    else
                    {
                        if (node.Name != node_name)
                            throw new DOMException(string.Format("{0}节点没有正常关闭", node.Name))
                            {
                                ColIndex = index
                            };
                        node_begin = false;
                        node.IsClosed = true;
                        node_outer_text.AppendFormat("</{0}>", node_name);
                        index += node_name.Length + 3;
                        break;
                    }
                }
                else if (text[index] == '<') //标签开始
                {
                    string node_name = get_next_token(text, ' ', index + 1);
                    if (node_name.Length == 0)
                        continue;
                    if (!node_begin)
                    {
                        node = new DOMNode();
                        node_begin = true;
                        node.Name = node_name;
                        node_outer_text.AppendFormat("<{0} ", node_name);
                        string node_remain = get_next_token(text, '>', index + node_name.Length + 2);
                        if (node_remain.Length == 0)
                            continue;
                        parse_node_attribute(node, node_remain);
                        node_outer_text.AppendFormat("{0}>", node_remain);
                        index += node_outer_text.Length;
                    }
                    else
                    {
                        if (node.Nodes == null)
                            node.Nodes = new List<DOMNode>();
                        DOMNode child_node = parse_next_node(text, ref index);
                        node.Nodes.Add(child_node);
                        node_inner_text.Append(child_node.OuterText);
                        node_outer_text.Append(child_node.OuterText);
                    }
                }
                else
                {
                    string node_text = get_next_token(text, '<', index);
                    node_outer_text.Append(node_text);
                    node_inner_text.Append(node_text);
                    if (!node_begin)
                    {
                        node = new DOMNode()
                            {
                                Name = "text",
                                IsClosed = true,
                            };
                        index += node_text.Length;
                        break;
                    }
                    else
                    {
                        index += node_text.Length;
                    }
                }
            }
            node.InnerText = node_inner_text.ToString();
            node.OuterText = node_outer_text.ToString();
            return node;
        }

        private void parse_node_attribute(DOMNode node, string text)
        {
            var matches = reg_att.Matches(text);
            if (matches.Count > 0)
                node.Attributes = new Dictionary<string, string>();
            foreach (Match match in matches)
            {
                node.Attributes.Add(match.Groups[1].Value, match.Groups[2].Value);
            }
        }

        string get_next_token(string text, char token_char, int index)
        {
            StringBuilder token = new StringBuilder();
            for (int i = index; i < text.Length; i++)
            {
                if (text[i] == token_char)
                    return token.ToString();
                token.Append(text[i]);
            }
            return token.ToString();
        }
    }

    public class DOMNode
    {
        public bool IsClosed { get; set; }
        public string Name { get; set; }
        public string InnerText { get; set; }
        public string OuterText { get; set; }
        public Dictionary<string, string> Attributes{ get; set; }
        public List<DOMNode> Nodes { get; set; }
    }

    //public class DOMAttribute
    //{
    //    public string Name { get; set; }
    //    public string Value { get; set; }
    //}

    public class DOMException : ApplicationException
    {
        public DOMException(string message)
            :base(message)
        {
        }
        public int RowIndex { get; set; }
        public int ColIndex { get; set; }
    }
}
