import React from "react";
import { BlockMath, InlineMath } from "react-katex";
import "katex/dist/katex.min.css"; // Import KaTeX CSS for proper styling

interface Choices {
    [key: string]: string;
}

interface JsonData {
    question: string;
    choices: Choices;
}

interface PromptRendererProps {
    prompt: string;
}

const PromptRenderer: React.FC<PromptRendererProps> = ({ prompt }) => {
    const renderContent = (content: string) => {
        try {
            const jsonData: JsonData = JSON.parse(content);
            if (jsonData.question && jsonData.choices) {
                return (
                    <div>
                        <h4>{jsonData.question}</h4>
                        {Object.entries(jsonData.choices).map(([key, value]) => (
                            <div key={key}>
                                <label>
                                    <input type="radio" name="choices" value={key} /> {key + ". " + value}
                                </label>
                            </div>
                        ))}
                    </div>
                );
            }
        } catch (e) {
            const parts = content.split(/(\$\$.+?\$\$|\$.+?\$)/g);
            return parts.map((part, index) => {
                if (part.startsWith("$$") && part.endsWith("$$")) {
                    // Render block-level LaTeX
                    const formula = part.slice(2, -2);
                    return <BlockMath key={index} math={formula} />;
                } else if (part.startsWith("$") && part.endsWith("$")) {
                    // Render inline-level LaTeX
                    const formula = part.slice(1, -1);
                    return <InlineMath key={index} math={formula} />;
                } else {
                    // Render regular text
                    return <span key={index}>{part}</span>;
                }
            });
        }
    };

    console.log("prompt: ", prompt);
    const parts = prompt
        .replace('{"', '<json>{"')
        .replace("}}", "}}<json>")
        .split(/<json>/);

    return (
        <div>
            {parts.map((part, index) => (
                <div key={index}>{renderContent(part)}</div>
            ))}
        </div>
    );
};

export default PromptRenderer;
