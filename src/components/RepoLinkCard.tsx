import { Card, CardHeader, CardTitle, CardDescription, CardContent, CardFooter } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Github, ExternalLink } from "lucide-react";

interface RepoLinkCardProps {
  href: string;
}

const RepoLinkCard = ({ href }: RepoLinkCardProps) => {
  return (
    <Card className="group overflow-hidden border ring-1 ring-transparent transition-all duration-300 hover:ring-ring">
      <CardHeader>
        <CardTitle className="text-2xl">Open on GitHub</CardTitle>
        <CardDescription>View the Calculator Fruit repository and code.</CardDescription>
      </CardHeader>
      <CardContent>
        <div className="flex items-center gap-3 text-muted-foreground">
          <Github className="h-6 w-6" aria-hidden="true" />
          <span className="break-all">{href}</span>
        </div>
      </CardContent>
      <CardFooter>
        <Button asChild size="lg">
          <a
            href={href}
            target="_blank"
            rel="noopener noreferrer"
            aria-label="Open Calculator Fruit repository on GitHub"
          >
            Visit Repository
            <ExternalLink className="ml-2 h-4 w-4" aria-hidden="true" />
          </a>
        </Button>
      </CardFooter>
    </Card>
  );
};

export default RepoLinkCard;
