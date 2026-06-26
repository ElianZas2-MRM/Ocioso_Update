import { useState } from "react";
import { AutomationScheduler } from "./components/AutomationScheduler";
import type { ScheduleConfig } from "./components/AutomationScheduler";
import { Toaster } from "./components/ui/sonner";
import { Button } from "./components/ui/button";
import { Badge } from "./components/ui/badge";
import {
  Settings2, Play, Square, Loader2, CheckCircle2,
  AlertCircle, Globe, Clock, Calendar, X,
} from "lucide-react";
import { toast } from "sonner";

type AutomationStatus = "idle" | "scheduled" | "running" | "completed" | "stopped";

const COUNTRIES_FLAG: Record<string, string> = {
  Argentina: "🇦🇷", Bolivia: "🇧🇴", Brasil: "🇧🇷",
  Chile: "🇨🇱", Colombia: "🇨🇴", Ecuador: "🇪🇨",
  Paraguay: "🇵🇾", Peru: "🇵🇪", Uruguay: "🇺🇾",
};

export default function App() {
  const [config, setConfig] = useState<ScheduleConfig | null>(null);
  const [status, setStatus] = useState<AutomationStatus>("idle");
  const [modalOpen, setModalOpen] = useState(false);

  const isLocked = status === "running" || status === "scheduled";
  const hasConfig = config !== null;

  const totalSlots = config
    ? Object.values(config.schedule).reduce((t, h) => t + h.length, 0)
    : 0;

  const handleSaveConfig = (newConfig: ScheduleConfig) => {
    setConfig(newConfig);
    setStatus("idle");
    setModalOpen(false);
  };

  const handleSchedule = () => {
    setStatus("scheduled");
    toast.success("Test automático programado", {
      description: "Se ejecutará según el horario configurado.",
    });
  };

  const handleDeactivate = () => {
    setStatus("idle");
    toast.info("Automatización desactivada");
  };

  const handleStart = () => {
    setStatus("running");
    toast.info("Test en ejecución...");
    setTimeout(() => {
      setStatus("completed");
      toast.success("Test completado", {
        description: "Todos los países fueron procesados correctamente.",
      });
    }, 4000);
  };

  const handleStop = () => {
    setStatus("stopped");
    toast.warning("Test detenido manualmente.");
  };

  const statusBadge = {
    idle:      { label: "Sin configurar",  cls: "bg-white/10 text-foreground border-white/25" },
    scheduled: { label: "Programado",      cls: "bg-primary/30 text-primary-foreground border-primary/50" },
    running:   { label: "Ejecutando...",   cls: "bg-amber-500/30 text-amber-200 border-amber-400/40" },
    completed: { label: "Completado",      cls: "bg-green-500/30 text-green-200 border-green-400/40" },
    stopped:   { label: "Detenido",        cls: "bg-red-500/30 text-red-200 border-red-400/40" },
  }[status];

  return (
    <div className="size-full min-h-screen bg-background flex items-center justify-center p-6">

      <div className="w-full max-w-2xl space-y-4">

        {/* Title row */}
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-lg font-bold text-foreground">Test Automático</h1>
            <p className="text-xs text-muted-foreground mt-0.5">Automatización de leads por país y horario</p>
          </div>
          <Badge className={`text-xs px-3 py-1.5 border flex items-center gap-1.5 ${statusBadge.cls}`}>
            {status === "running" && <Loader2 className="size-3.5 animate-spin" />}
            {status === "scheduled" && <Calendar className="size-3.5" />}
            {status === "completed" && <CheckCircle2 className="size-3.5" />}
            {status === "stopped" && <Square className="size-3.5" />}
            {status === "idle" && <Clock className="size-3.5" />}
            {statusBadge.label}
          </Badge>
        </div>

        {/* Main panel */}
        <div className="rounded-2xl border border-white/20 bg-card shadow-2xl overflow-hidden">

          {/* Summary or empty state */}
          <div className="p-6">
            {hasConfig ? (
              <div className="space-y-4">
                <div>
                  <p className="text-xs uppercase tracking-widest text-muted-foreground font-semibold mb-2 flex items-center gap-1.5">
                    <Globe className="size-3.5" /> Países a testear
                  </p>
                  <div className="flex flex-wrap gap-2">
                    {config.countries.map((c) => (
                      <span key={c} className="flex items-center gap-1.5 text-sm px-2.5 py-1 rounded-lg bg-white/10 border border-white/20 font-medium">
                        <span>{COUNTRIES_FLAG[c]}</span> {c}
                      </span>
                    ))}
                  </div>
                </div>
                <div>
                  <p className="text-xs uppercase tracking-widest text-muted-foreground font-semibold mb-2 flex items-center gap-1.5">
                    <Clock className="size-3.5" /> Horario configurado
                    <span className="text-primary font-bold">({totalSlots} slot{totalSlots !== 1 ? "s" : ""})</span>
                  </p>
                  <div className="space-y-1.5">
                    {Object.entries(config.schedule)
                      .filter(([_, h]) => h.length > 0)
                      .map(([day, hours]) => (
                        <div key={day} className="flex items-center gap-3 text-sm">
                          <span className="font-semibold min-w-[84px] text-foreground">{day}:</span>
                          <div className="flex flex-wrap gap-1">
                            {hours.map((h) => (
                              <span key={h} className="text-xs px-2 py-0.5 rounded bg-primary/25 border border-primary/40 text-primary-foreground font-medium">{h}</span>
                            ))}
                          </div>
                        </div>
                      ))}
                  </div>
                </div>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center py-10 text-center gap-3">
                <div className="size-14 rounded-full bg-white/10 border border-white/20 flex items-center justify-center">
                  <Settings2 className="size-6 text-muted-foreground" />
                </div>
                <div>
                  <p className="text-sm font-semibold text-foreground">Sin configuración</p>
                  <p className="text-xs text-muted-foreground mt-1">Configurá los días, horarios y países para poder programar el test.</p>
                </div>
              </div>
            )}
          </div>

          {/* Progress / status strips */}
          {status === "running" && (
            <div className="px-6 pb-4">
              <div className="h-1.5 rounded-full bg-white/10 overflow-hidden">
                <div className="h-full bg-amber-400 rounded-full animate-pulse w-2/3" />
              </div>
              <p className="text-xs text-amber-300 mt-1.5 flex items-center gap-1.5">
                <Loader2 className="size-3 animate-spin" /> Procesando países...
              </p>
            </div>
          )}
          {status === "completed" && (
            <div className="px-6 pb-4">
              <div className="h-1.5 rounded-full bg-green-500/30 overflow-hidden">
                <div className="h-full bg-green-400 rounded-full w-full" />
              </div>
              <p className="text-xs text-green-300 mt-1.5 flex items-center gap-1.5">
                <CheckCircle2 className="size-3" /> Todos los países procesados correctamente.
              </p>
            </div>
          )}
          {status === "stopped" && (
            <div className="px-6 pb-4">
              <p className="text-xs text-red-300 flex items-center gap-1.5">
                <AlertCircle className="size-3" /> El test fue detenido manualmente.
              </p>
            </div>
          )}

          <div className="border-t border-white/10" />

          {/* Actions */}
          <div className="px-6 py-4 flex items-center justify-center gap-3 flex-wrap">
            <Button onClick={() => setModalOpen(true)} disabled={isLocked} variant="secondary"
              className="gap-2 font-bold uppercase tracking-wide text-xs">
              <Settings2 className="size-4" />
              {hasConfig ? "Editar configuración" : "Configurar automatización"}
            </Button>

            {!isLocked && (status === "idle" || status === "completed" || status === "stopped") && (
              status === "idle" ? (
                <Button onClick={handleSchedule} disabled={!hasConfig}
                  className="gap-2 font-bold uppercase tracking-wide text-xs shadow-lg shadow-primary/30">
                  <Play className="size-4" />
                  Programar test automático
                </Button>
              ) : (
                <Button onClick={handleDeactivate} variant="secondary"
                  className="gap-2 font-bold uppercase tracking-wide text-xs border-red-400/40 text-red-300 hover:bg-red-500/20">
                  <Square className="size-4" />
                  Desactivar
                </Button>
              )
            )}

            {status === "scheduled" && (
              <>
                <Button onClick={handleStart} size="sm"
                  className="gap-2 bg-amber-500 hover:bg-amber-400 text-white border-0 font-bold uppercase tracking-wide text-xs">
                  <Play className="size-3.5" />
                  Iniciar ahora
                </Button>
                <Button onClick={handleDeactivate} variant="secondary" size="sm"
                  className="gap-2 font-bold uppercase tracking-wide text-xs border-red-400/40 text-red-300 hover:bg-red-500/20">
                  <Square className="size-3.5" />
                  Desactivar
                </Button>
              </>
            )}

            {status === "running" && (
              <Button onClick={handleStop}
                className="gap-2 bg-red-500 hover:bg-red-400 text-white border-0 font-bold uppercase tracking-wide text-xs">
                <Square className="size-4" />
                Detener
              </Button>
            )}
          </div>
        </div>
      </div>

      {/* Custom modal — no Radix Dialog to avoid ref/portal issues */}
      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
          {/* Backdrop */}
          <div
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setModalOpen(false)}
          />
          {/* Panel */}
          <div className="relative z-10 w-full max-w-3xl bg-card border border-white/20 rounded-2xl shadow-2xl flex flex-col max-h-[90vh]">
            {/* Modal header */}
            <div className="flex items-center justify-between px-6 py-4 border-b border-white/10 shrink-0">
              <div className="flex items-center gap-2">
                <Settings2 className="size-5 text-primary" />
                <span className="text-lg font-semibold">Configurar automatización</span>
              </div>
              <button
                onClick={() => setModalOpen(false)}
                className="rounded-full p-1.5 hover:bg-white/10 transition-colors text-muted-foreground hover:text-foreground"
              >
                <X className="size-5" />
              </button>
            </div>
            {/* Modal body */}
            <div className="overflow-y-auto p-6">
              <AutomationScheduler
                onSave={handleSaveConfig}
                initialConfig={config ?? undefined}
              />
            </div>
          </div>
        </div>
      )}

      <Toaster />
    </div>
  );
}
